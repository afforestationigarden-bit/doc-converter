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

const C = { ink: '#12343B', paper: '#F7F5F0', accent: '#E0A458', line: '#D9D4C7', white: '#FFFFFF' };

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

const pageCss = `<style>body{font-family:sans-serif;font-size:12pt;line-height:1.5;padding:24px;color:#111}
img{max-width:100%;max-height:95vh;display:block;margin:0 auto;page-break-after:always}
h1,h2,h3{margin:12px 0 6px}p{margin:0 0 8px}</style>`;

async function readBase64(uri) {
  return FileSystem.readAsStringAsync(uri, { encoding: 'base64' });
}

async function docxToHtml(uri) {
  const zip = await JSZip.loadAsync(await readBase64(uri), { base64: true });
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('এটি সঠিক .docx ফাইল নয়');
  const xml = await xmlFile.async('string');
  const paras = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  const out = paras.map((p) => {
    const style = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || '';
    const tag = /^Heading1|^Title/i.test(style) ? 'h1' : /^Heading2/i.test(style) ? 'h2' : /^Heading3/i.test(style) ? 'h3' : 'p';
    const runs = p.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
    const html = runs.map((r) => {
      let t = '';
      const parts = r.match(/<w:t[^>]*>[\s\S]*?<\/w:t>|<w:tab\/>|<w:br\/>/g) || [];
      parts.forEach((x) => {
        if (x === '<w:tab/>') t += '&emsp;';
        else if (x === '<w:br/>') t += '<br/>';
        else t += esc(decode(x.replace(/<[^>]+>/g, '')));
      });
      if (/<w:b\/>|<w:b /.test(r)) t = `<b>${t}</b>`;
      if (/<w:i\/>|<w:i /.test(r)) t = `<i>${t}</i>`;
      return t;
    }).join('');
    return `<${tag}>${html || '&nbsp;'}</${tag}>`;
  });
  return pageCss + out.join('');
}

async function savePdf(html, name) {
  const { uri } = await Print.printToFileAsync({ html });
  const dest = FileSystem.documentDirectory + name + '.pdf';
  await FileSystem.copyAsync({ from: uri, to: dest });
  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(dest, { mimeType: 'application/pdf', dialogTitle: 'PDF সেভ বা শেয়ার করো' });
  } else {
    Alert.alert('তৈরি হয়েছে', dest);
  }
}

const TOOLS = [
  { id: 'img', title: 'ছবি → PDF', hint: 'এক বা একাধিক ছবি বেছে একটি PDF বানাও' },
  { id: 'txt', title: 'লেখা → PDF', hint: 'টাইপ বা পেস্ট করা লেখা থেকে PDF' },
  { id: 'docx', title: 'Word → PDF', hint: '.docx ফাইলের লেখা ও মৌলিক ফরম্যাট থেকে PDF' },
];

export default function App() {
  const [tool, setTool] = useState(null);
  const [text, setText] = useState('');
  const [busy, setBusy] = useState(false);

  const run = async (fn) => {
    setBusy(true);
    try { await fn(); } catch (e) { Alert.alert('সমস্যা হয়েছে', String(e.message || e)); }
    setBusy(false);
  };

  const convertImages = () => run(async () => {
    const res = await DocumentPicker.getDocumentAsync({ type: 'image/*', multiple: true, copyToCacheDirectory: true });
    if (res.canceled) return;
    const imgs = [];
    for (const a of res.assets) {
      const mime = a.mimeType || 'image/jpeg';
      imgs.push(`<img src="data:${mime};base64,${await readBase64(a.uri)}"/>`);
    }
    await savePdf(pageCss + imgs.join(''), 'images-' + Date.now());
  });

  const convertText = () => run(async () => {
    if (!text.trim()) throw new Error('আগে কিছু লেখো');
    const body = text.split('\n').map((l) => `<p>${esc(l) || '&nbsp;'}</p>`).join('');
    await savePdf(pageCss + body, 'text-' + Date.now());
  });

  const convertDocx = () => run(async () => {
    const res = await DocumentPicker.getDocumentAsync({
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      copyToCacheDirectory: true,
    });
    if (res.canceled) return;
    const a = res.assets[0];
    await savePdf(await docxToHtml(a.uri), (a.name || 'document').replace(/\.docx$/i, ''));
  });

  const go = { img: convertImages, txt: convertText, docx: convertDocx };

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

        {tool === 'txt' && (
          <TextInput style={s.input} multiline placeholder="এখানে লেখো বা পেস্ট করো" value={text} onChangeText={setText} textAlignVertical="top" />
        )}

        {tool && (
          <TouchableOpacity style={s.btn} disabled={busy} onPress={go[tool]}>
            {busy ? <ActivityIndicator color={C.ink} /> : (
              <Text style={s.btnText}>{tool === 'txt' ? 'PDF বানাও' : 'ফাইল বেছে PDF বানাও'}</Text>
            )}
          </TouchableOpacity>
        )}

        <Text style={s.note}>PDF → Word এই সংস্করণে নেই। Word ফাইলের ছবি ও জটিল টেবিল PDF-এ আসবে না।</Text>
      </ScrollView>
    </SafeAreaView>
  );
}

const s = StyleSheet.create({
  root: { flex: 1, backgroundColor: C.paper },
  scroll: { padding: 20, paddingTop: 36 },
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
