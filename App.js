import React, { useState } from 'react';
import {
  View, Text, TextInput, TouchableOpacity, ScrollView,
  ActivityIndicator, Alert, StyleSheet, SafeAreaView,
} from 'react-native';
import { StatusBar } from 'expo-status-bar';
import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system/legacy';
import * as Print from 'expo-print';
import * as Sharing from 'expo-sharing';
import JSZip from 'jszip';
import KonvertPdf from './modules/konvert-pdf';

const C = { ink: '#12343B', paper: '#F7F5F0', accent: '#E0A458', line: '#D9D4C7', white: '#FFFFFF' };
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

const pageCss = `<style>body{font-family:sans-serif;font-size:12pt;line-height:1.5;padding:24px;color:#111}
img{max-width:100%;max-height:95vh;display:block;margin:0 auto;page-break-after:always}
h1,h2,h3{margin:12px 0 6px}p{margin:0 0 8px}</style>`;

const OUT = FileSystem.cacheDirectory + 'konvert-out/';
const LEVELS = {
  high: { label: 'High', hint: 'Sharpest, smaller reduction', maxDim: 2400, quality: 85 },
  medium: { label: 'Medium', hint: 'Balanced', maxDim: 1800, quality: 70 },
  low: { label: 'Low', hint: 'Smallest file', maxDim: 1300, quality: 55 },
};
const fmtSize = (b) => (b >= 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB');

const readBase64 = (uri) => FileSystem.readAsStringAsync(uri, { encoding: 'base64' });

async function makeFile(name, data, encoding, mime) {
  const dest = FileSystem.documentDirectory + name;
  await FileSystem.writeAsStringAsync(dest, data, { encoding });
  return { uri: dest, name, mime };
}

async function savePdfFromHtml(html, name) {
  const { uri } = await Print.printToFileAsync({ html });
  const fileName = name + '.pdf';
  const dest = FileSystem.documentDirectory + fileName;
  await FileSystem.copyAsync({ from: uri, to: dest });
  return { uri: dest, name: fileName, mime: 'application/pdf' };
}

async function docxParagraphs(uri) {
  const zip = await JSZip.loadAsync(await readBase64(uri), { base64: true });
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('This is not a valid .docx file');
  const xml = await xmlFile.async('string');
  const paras = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  return paras.map((p) => {
    const style = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || '';
    const tag = /^Heading1|^Title/i.test(style) ? 'h1' : /^Heading2/i.test(style) ? 'h2' : /^Heading3/i.test(style) ? 'h3' : 'p';
    const runs = p.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
    let html = '';
    let text = '';
    runs.forEach((r) => {
      let h = '';
      let t = '';
      const parts = r.match(/<w:t[^>]*>[\s\S]*?<\/w:t>|<w:tab\/>|<w:br\/>/g) || [];
      parts.forEach((x) => {
        if (x === '<w:tab/>') { h += '&emsp;'; t += '\t'; }
        else if (x === '<w:br/>') { h += '<br/>'; t += '\n'; }
        else { const raw = decode(x.replace(/<[^>]+>/g, '')); h += esc(raw); t += raw; }
      });
      if (/<w:b\/>|<w:b /.test(r)) h = `<b>${h}</b>`;
      if (/<w:i\/>|<w:i /.test(r)) h = `<i>${h}</i>`;
      html += h;
      text += t;
    });
    return { tag, html, text };
  });
}

async function buildDocx(text) {
  const body = text.split('\n').map((l) =>
    `<w:p><w:r><w:t xml:space="preserve">${esc(l)}</w:t></w:r></w:p>`).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: 'base64' });
}

function parseRanges(input, total) {
  const idx = new Set();
  input.split(',').map((x) => x.trim()).filter(Boolean).forEach((part) => {
    const m = part.match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!m) throw new Error(`"${part}" is not valid. Example: 1-3,5`);
    const a = parseInt(m[1], 10);
    const b = m[2] ? parseInt(m[2], 10) : a;
    if (a < 1 || b < a || b > total) throw new Error(`Page numbers must be between 1 and ${total}`);
    for (let i = a; i <= b; i++) idx.add(i - 1);
  });
  if (!idx.size) throw new Error('Enter which pages to extract');
  return [...idx];
}

const TOOLS = [
  { id: 'img', title: 'Image → PDF', hint: 'Combine one or more images into a PDF' },
  { id: 'txt', title: 'Text → PDF', hint: 'Turn typed or pasted text into a PDF' },
  { id: 'docx', title: 'Word → PDF', hint: 'Convert a .docx file (text and basic formatting) to PDF' },
  { id: 'docx2txt', title: 'Word → Text (.txt)', hint: 'Extract plain text from a .docx file' },
  { id: 'txt2docx', title: 'Text → Word (.docx)', hint: 'Turn typed or pasted text into a Word file' },
  { id: 'merge', title: 'Merge PDFs', hint: 'Combine multiple PDFs into one (works with large files)' },
  { id: 'split', title: 'Split PDF', hint: 'Extract chosen pages into a new PDF (works with large files)' },
  { id: 'compress', title: 'Compress PDF', hint: 'Reduce file size while keeping text and images clear' },
];
const NEEDS_TEXT = ['txt', 'txt2docx'];

export default function App() {
  const [tool, setTool] = useState(null);
  const [text, setText] = useState('');
  const [range, setRange] = useState('');
  const [level, setLevel] = useState('medium');
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState(null);
  const [dirUri, setDirUri] = useState(null);

  const run = async (fn) => {
    setBusy(true);
    try {
      const r = await fn();
      if (r) {
        if (result && result.uri !== r.uri) FileSystem.deleteAsync(result.uri, { idempotent: true }).catch(() => {});
        setResult(r);
      }
    } catch (e) { Alert.alert('Something went wrong', String(e.message || e)); }
    setBusy(false);
  };

  const shareFile = () => run(async () => {
    if (!result) return;
    if (!(await Sharing.isAvailableAsync())) throw new Error('Sharing is not available on this phone');
    await Sharing.shareAsync(result.uri, { mimeType: result.mime, dialogTitle: 'Share' });
  });

  const downloadFile = () => run(async () => {
    if (!result) return;
    const SAF = FileSystem.StorageAccessFramework;
    let dir = dirUri;
    if (!dir) {
      const perm = await SAF.requestDirectoryPermissionsAsync();
      if (!perm.granted) return;
      dir = perm.directoryUri;
      setDirUri(dir);
    }
    const dest = await SAF.createFileAsync(dir, result.name.replace(/\.[^.]+$/, ''), result.mime);
    await KonvertPdf.copyToUri(result.uri, dest);
    Alert.alert('Downloaded', result.name + ' was saved to the folder you chose.');
  });

  const pick = async (type, multiple = false) => {
    const res = await DocumentPicker.getDocumentAsync({ type, multiple, copyToCacheDirectory: true });
    return res.canceled ? null : res.assets;
  };

  const actions = {
    img: async () => {
      const files = await pick('image/*', true);
      if (!files) return;
      const imgs = [];
      for (const a of files) {
        imgs.push(`<img src="data:${a.mimeType || 'image/jpeg'};base64,${await readBase64(a.uri)}"/>`);
      }
      return savePdfFromHtml(pageCss + imgs.join(''), 'images-' + Date.now());
    },
    txt: async () => {
      if (!text.trim()) throw new Error('Type some text first');
      const body = text.split('\n').map((l) => `<p>${esc(l) || '&nbsp;'}</p>`).join('');
      return savePdfFromHtml(pageCss + body, 'text-' + Date.now());
    },
    docx: async () => {
      const f = await pick(DOCX);
      if (!f) return;
      const paras = await docxParagraphs(f[0].uri);
      const html = pageCss + paras.map((p) => `<${p.tag}>${p.html || '&nbsp;'}</${p.tag}>`).join('');
      return savePdfFromHtml(html, (f[0].name || 'document').replace(/\.docx$/i, ''));
    },
    docx2txt: async () => {
      const f = await pick(DOCX);
      if (!f) return;
      const paras = await docxParagraphs(f[0].uri);
      const name = (f[0].name || 'document').replace(/\.docx$/i, '') + '.txt';
      return makeFile(name, paras.map((p) => p.text).join('\n'), 'utf8', 'text/plain');
    },
    txt2docx: async () => {
      if (!text.trim()) throw new Error('Type some text first');
      return makeFile('text-' + Date.now() + '.docx', await buildDocx(text), 'base64', DOCX);
    },
    merge: async () => {
      const files = await pick('application/pdf', true);
      if (!files) return;
      if (files.length < 2) throw new Error('Select at least 2 PDFs');
      const name = 'merged-' + Date.now() + '.pdf';
      const r = await KonvertPdf.mergePdfs(files.map((f) => f.uri), OUT + name);
      return { uri: 'file://' + r.path, name, mime: 'application/pdf', info: 'Size: ' + fmtSize(r.outputBytes) };
    },
    split: async () => {
      if (!range.trim()) throw new Error('Enter page numbers first, e.g. 1-3,5');
      const f = await pick('application/pdf');
      if (!f) return;
      const total = await KonvertPdf.getPageCount(f[0].uri);
      const idx = parseRanges(range, total);
      const name = 'pages-' + Date.now() + '.pdf';
      const r = await KonvertPdf.splitPdf(f[0].uri, idx, OUT + name);
      return { uri: 'file://' + r.path, name, mime: 'application/pdf', info: idx.length + ' of ' + total + ' pages, ' + fmtSize(r.outputBytes) };
    },
    compress: async () => {
      const f = await pick('application/pdf');
      if (!f) return;
      const p = LEVELS[level];
      const name = (f[0].name || 'document').replace(/\.pdf$/i, '') + '-compressed.pdf';
      const r = await KonvertPdf.compressPdf(f[0].uri, OUT + name, p.maxDim, p.quality);
      const before = f[0].size || 0;
      let info = before > 0
        ? fmtSize(before) + ' → ' + fmtSize(r.outputBytes) + ' (' + Math.round((1 - r.outputBytes / before) * 100) + '% smaller)'
        : 'Size: ' + fmtSize(r.outputBytes);
      info += '\n' + r.replaced + ' of ' + r.total + ' images were recompressed.';
      return { uri: 'file://' + r.path, name, mime: 'application/pdf', info };
    },
  };

  return (
    <SafeAreaView style={s.root}>
      <StatusBar style="dark" />
      <ScrollView contentContainerStyle={s.scroll} keyboardShouldPersistTaps="handled">
        <Text style={s.title}>Konvert</Text>
        <Text style={s.sub}>Convert files without internet</Text>

        {TOOLS.map((t) => (
          <TouchableOpacity key={t.id} style={[s.card, tool === t.id && s.cardOn]} onPress={() => { setTool(t.id); setResult(null); }}>
            <Text style={s.cardTitle}>{t.title}</Text>
            <Text style={s.cardHint}>{t.hint}</Text>
          </TouchableOpacity>
        ))}

        {NEEDS_TEXT.includes(tool) && (
          <TextInput style={s.input} multiline placeholder="Type or paste your text here" value={text} onChangeText={setText} textAlignVertical="top" />
        )}
        {tool === 'split' && (
          <TextInput style={[s.input, { minHeight: 52 }]} placeholder="Which pages? e.g. 1-3,5" value={range} onChangeText={setRange} keyboardType="numbers-and-punctuation" />
        )}

        {tool === 'compress' && (
          <View style={s.levels}>
            {Object.keys(LEVELS).map((k) => (
              <TouchableOpacity key={k} style={[s.level, level === k && s.levelOn]} onPress={() => setLevel(k)}>
                <Text style={s.levelTitle}>{LEVELS[k].label}</Text>
                <Text style={s.levelHint}>{LEVELS[k].hint}</Text>
              </TouchableOpacity>
            ))}
          </View>
        )}

        {tool && (
          <TouchableOpacity style={s.btn} disabled={busy} onPress={() => run(actions[tool])}>
            {busy ? <ActivityIndicator color={C.ink} /> : (
              <Text style={s.btnText}>{NEEDS_TEXT.includes(tool) ? 'Create' : 'Choose file & start'}</Text>
            )}
          </TouchableOpacity>
        )}

        {busy && <Text style={s.note}>Working… large files can take a few minutes. Please keep the app open.</Text>}

        {result && (
          <View style={s.result}>
            <Text style={s.resultTitle}>File created</Text>
            <Text style={s.resultName}>{result.name}</Text>
            {result.info ? <Text style={s.resultInfo}>{result.info}</Text> : null}
            <View style={s.row}>
              <TouchableOpacity style={[s.btn, s.half]} disabled={busy} onPress={downloadFile}>
                <Text style={s.btnText}>Download</Text>
              </TouchableOpacity>
              <TouchableOpacity style={[s.btn, s.half, s.btnAlt]} disabled={busy} onPress={shareFile}>
                <Text style={s.btnText}>Share</Text>
              </TouchableOpacity>
            </View>
          </View>
        )}

        <Text style={s.note}>PDF → Word and PDF → Text are not in this version. Images and complex tables in Word files will not appear in the PDF.</Text>
      </ScrollView>
    </SafeAreaView>
  );
}

const s = StyleSheet.create({
  root: { flex: 1, backgroundColor: C.paper },
  scroll: { padding: 20, paddingTop: 36, paddingBottom: 60 },
  title: { fontSize: 30, fontWeight: '700', color: C.ink },
  sub: { fontSize: 15, color: C.ink, opacity: 0.7, marginBottom: 22, marginTop: 4 },
  card: { backgroundColor: C.white, borderColor: C.line, borderWidth: 1, borderRadius: 14, padding: 16, marginBottom: 12 },
  cardOn: { borderColor: C.ink, borderWidth: 2 },
  cardTitle: { fontSize: 18, fontWeight: '600', color: C.ink },
  cardHint: { fontSize: 14, color: C.ink, opacity: 0.65, marginTop: 4 },
  input: { backgroundColor: C.white, borderColor: C.line, borderWidth: 1, borderRadius: 14, padding: 14, minHeight: 160, fontSize: 16, marginBottom: 12 },
  btn: { backgroundColor: C.accent, borderRadius: 14, padding: 16, alignItems: 'center', marginTop: 4 },
  btnText: { fontSize: 17, fontWeight: '700', color: C.ink },
  result: { backgroundColor: C.white, borderColor: C.ink, borderWidth: 2, borderRadius: 14, padding: 16, marginTop: 16 },
  resultTitle: { fontSize: 16, fontWeight: '700', color: C.ink },
  resultName: { fontSize: 14, color: C.ink, opacity: 0.7, marginTop: 4, marginBottom: 12 },
  resultInfo: { fontSize: 14, color: C.ink, marginBottom: 12, lineHeight: 20 },
  levels: { flexDirection: 'row', gap: 8, marginBottom: 12 },
  level: { flex: 1, backgroundColor: C.white, borderColor: C.line, borderWidth: 1, borderRadius: 12, padding: 10 },
  levelOn: { borderColor: C.ink, borderWidth: 2 },
  levelTitle: { fontSize: 15, fontWeight: '700', color: C.ink },
  levelHint: { fontSize: 12, color: C.ink, opacity: 0.65, marginTop: 2 },
  row: { flexDirection: 'row', gap: 10 },
  half: { flex: 1, marginTop: 0 },
  btnAlt: { backgroundColor: C.paper, borderColor: C.ink, borderWidth: 1.5 },
  note: { fontSize: 13, color: C.ink, opacity: 0.6, marginTop: 20, lineHeight: 19 },
});
