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
import { PDFDocument } from 'pdf-lib';

const C = { ink: '#12343B', paper: '#F7F5F0', accent: '#E0A458', line: '#D9D4C7', white: '#FFFFFF' };
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

const pageCss = `<style>body{font-family:sans-serif;font-size:12pt;line-height:1.5;padding:24px;color:#111}
img{max-width:100%;max-height:95vh;display:block;margin:0 auto;page-break-after:always}
h1,h2,h3{margin:12px 0 6px}p{margin:0 0 8px}</style>`;

const readBase64 = (uri) => FileSystem.readAsStringAsync(uri, { encoding: 'base64' });

async function shareData(name, data, encoding, mime) {
  const dest = FileSystem.documentDirectory + name;
  await FileSystem.writeAsStringAsync(dest, data, { encoding });
  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(dest, { mimeType: mime, dialogTitle: 'সেভ বা শেয়ার করো' });
  } else {
    Alert.alert('তৈরি হয়েছে', dest);
  }
}

async function savePdfFromHtml(html, name) {
  const { uri } = await Print.printToFileAsync({ html });
  const dest = FileSystem.documentDirectory + name + '.pdf';
  await FileSystem.copyAsync({ from: uri, to: dest });
  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(dest, { mimeType: 'application/pdf', dialogTitle: 'PDF সেভ বা শেয়ার করো' });
  } else {
    Alert.alert('তৈরি হয়েছে', dest);
  }
}

async function docxParagraphs(uri) {
  const zip = await JSZip.loadAsync(await readBase64(uri), { base64: true });
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('এটি সঠিক .docx ফাইল নয়');
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
    if (!m) throw new Error(`"${part}" সঠিক নয়। উদাহরণ: 1-3,5`);
    const a = parseInt(m[1], 10);
    const b = m[2] ? parseInt(m[2], 10) : a;
    if (a < 1 || b < a || b > total) throw new Error(`পাতা ১ থেকে ${total}-এর মধ্যে হতে হবে`);
    for (let i = a; i <= b; i++) idx.add(i - 1);
  });
  if (!idx.size) throw new Error('কোন পাতা নেবে তা লেখো');
  return [...idx];
}

const TOOLS = [
  { id: 'img', title: 'ছবি → PDF', hint: 'এক বা একাধিক ছবি বেছে একটি PDF বানাও' },
  { id: 'txt', title: 'লেখা → PDF', hint: 'টাইপ বা পেস্ট করা লেখা থেকে PDF' },
  { id: 'docx', title: 'Word → PDF', hint: '.docx ফাইলের লেখা ও মৌলিক ফরম্যাট থেকে PDF' },
  { id: 'docx2txt', title: 'Word → লেখা (.txt)', hint: '.docx ফাইল থেকে সাধারণ লেখা বের করো' },
  { id: 'txt2docx', title: 'লেখা → Word (.docx)', hint: 'টাইপ বা পেস্ট করা লেখা থেকে Word ফাইল' },
  { id: 'merge', title: 'PDF জোড়া লাগাও', hint: 'একাধিক PDF বেছে একটি PDF বানাও' },
  { id: 'split', title: 'PDF থেকে পাতা আলাদা করো', hint: 'নির্দিষ্ট পাতা নিয়ে নতুন PDF বানাও' },
];
const NEEDS_TEXT = ['txt', 'txt2docx'];

export default function App() {
  const [tool, setTool] = useState(null);
  const [text, setText] = useState('');
  const [range, setRange] = useState('');
  const [busy, setBusy] = useState(false);

  const run = async (fn) => {
    setBusy(true);
    try { await fn(); } catch (e) { Alert.alert('সমস্যা হয়েছে', String(e.message || e)); }
    setBusy(false);
  };

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
      await savePdfFromHtml(pageCss + imgs.join(''), 'images-' + Date.now());
    },
    txt: async () => {
      if (!text.trim()) throw new Error('আগে কিছু লেখো');
      const body = text.split('\n').map((l) => `<p>${esc(l) || '&nbsp;'}</p>`).join('');
      await savePdfFromHtml(pageCss + body, 'text-' + Date.now());
    },
    docx: async () => {
      const f = await pick(DOCX);
      if (!f) return;
      const paras = await docxParagraphs(f[0].uri);
      const html = pageCss + paras.map((p) => `<${p.tag}>${p.html || '&nbsp;'}</${p.tag}>`).join('');
      await savePdfFromHtml(html, (f[0].name || 'document').replace(/\.docx$/i, ''));
    },
    docx2txt: async () => {
      const f = await pick(DOCX);
      if (!f) return;
      const paras = await docxParagraphs(f[0].uri);
      const name = (f[0].name || 'document').replace(/\.docx$/i, '') + '.txt';
      await shareData(name, paras.map((p) => p.text).join('\n'), 'utf8', 'text/plain');
    },
    txt2docx: async () => {
      if (!text.trim()) throw new Error('আগে কিছু লেখো');
      await shareData('text-' + Date.now() + '.docx', await buildDocx(text), 'base64', DOCX);
    },
    merge: async () => {
      const files = await pick('application/pdf', true);
      if (!files) return;
      if (files.length < 2) throw new Error('কমপক্ষে ২টি PDF বেছে নাও');
      const out = await PDFDocument.create();
      for (const a of files) {
        const src = await PDFDocument.load(await readBase64(a.uri));
        const pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach((p) => out.addPage(p));
      }
      await shareData('merged-' + Date.now() + '.pdf', await out.saveAsBase64(), 'base64', 'application/pdf');
    },
    split: async () => {
      if (!range.trim()) throw new Error('আগে পাতার নম্বর লেখো, যেমন 1-3,5');
      const f = await pick('application/pdf');
      if (!f) return;
      const src = await PDFDocument.load(await readBase64(f[0].uri));
      const idx = parseRanges(range, src.getPageCount());
      const out = await PDFDocument.create();
      (await out.copyPages(src, idx)).forEach((p) => out.addPage(p));
      await shareData('pages-' + Date.now() + '.pdf', await out.saveAsBase64(), 'base64', 'application/pdf');
    },
  };

  return (
    <SafeAreaView style={s.root}>
      <StatusBar style="dark" />
      <ScrollView contentContainerStyle={s.scroll} keyboardShouldPersistTaps="handled">
        <Text style={s.title}>Doc Converter</Text>
        <Text style={s.sub}>ইন্টারনেট ছাড়াই ফাইল কনভার্ট করো</Text>

        {TOOLS.map((t) => (
          <TouchableOpacity key={t.id} style={[s.card, tool === t.id && s.cardOn]} onPress={() => setTool(t.id)}>
            <Text style={s.cardTitle}>{t.title}</Text>
            <Text style={s.cardHint}>{t.hint}</Text>
          </TouchableOpacity>
        ))}

        {NEEDS_TEXT.includes(tool) && (
          <TextInput style={s.input} multiline placeholder="এখানে লেখো বা পেস্ট করো" value={text} onChangeText={setText} textAlignVertical="top" />
        )}
        {tool === 'split' && (
          <TextInput style={[s.input, { minHeight: 52 }]} placeholder="কোন পাতা? যেমন 1-3,5" value={range} onChangeText={setRange} keyboardType="numbers-and-punctuation" />
        )}

        {tool && (
          <TouchableOpacity style={s.btn} disabled={busy} onPress={() => run(actions[tool])}>
            {busy ? <ActivityIndicator color={C.ink} /> : (
              <Text style={s.btnText}>{NEEDS_TEXT.includes(tool) ? 'তৈরি করো' : 'ফাইল বেছে শুরু করো'}</Text>
            )}
          </TouchableOpacity>
        )}

        <Text style={s.note}>PDF → Word ও PDF → লেখা এই সংস্করণে নেই। Word ফাইলের ছবি ও জটিল টেবিল আসবে না।</Text>
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
  note: { fontSize: 13, color: C.ink, opacity: 0.6, marginTop: 20, lineHeight: 19 },
});
