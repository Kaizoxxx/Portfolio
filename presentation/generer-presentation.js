// Génère la présentation PowerPoint du portfolio BTS SIO SISR
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const fi = require('react-icons/fi');

const REPO = path.join(__dirname, '..');
const OUT = process.argv[2] || path.join(__dirname, 'Presentation-BTS-SIO-SISR-Kylian-Grafeille.pptx');

// Palette reprise du site
const C = {
  bg: '0D1117', surface: '161B22', surface2: '1C232D', border: '30363D',
  text: 'E6EDF3', muted: '8B949E', green: '3FB950', blue: '58A6FF',
  todoBg: '3B2E05', todo: 'E3B341', white: 'FFFFFF',
};
const MONO = 'Courier New';
const SANS = 'Calibri';

async function png(svg, size = 256) {
  const buf = await sharp(Buffer.from(svg)).resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}
async function featherIcon(name, color) {
  const svg = renderToStaticMarkup(React.createElement(fi[name], { color: '#' + color, size: 256, strokeWidth: 1.8 }));
  return png(svg);
}
async function brandIcon(file) {
  let svg = fs.readFileSync(path.join(REPO, 'assets/icons', file), 'utf8');
  svg = svg.replace(/currentColor/g, '#FFFFFF').replace('<svg ', '<svg fill="#FFFFFF" ');
  return png(svg);
}

(async () => {
  const I = {};
  const fIcons = {
    globe: 'FiGlobe', monitor: 'FiMonitor', headphones: 'FiHeadphones', user: 'FiUser', shield: 'FiShield',
    grid: 'FiGrid', users: 'FiUsers', check: 'FiCheckCircle', server: 'FiServer', map: 'FiMapPin',
    rss: 'FiRss', book: 'FiBookOpen', briefcase: 'FiBriefcase', target: 'FiTarget', wifi: 'FiWifi',
    award: 'FiAward', camera: 'FiCamera', trending: 'FiTrendingUp', alert: 'FiAlertTriangle', zap: 'FiZap', file: 'FiFileText',
  };
  for (const [k, v] of Object.entries(fIcons)) I[k] = await featherIcon(v, C.green);
  for (const [k, v] of Object.entries({ blueRss: 'FiRss', blueFile: 'FiFileText' })) I[k] = await featherIcon(v, C.blue);
  for (const k of ['kalilinux', 'windows', 'cisco', 'vmware', 'github', 'linkedin']) I[k] = await brandIcon(k + '.svg');

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
  pres.title = 'Portfolio BTS SIO SISR – Kylian Grafeille Clément';
  pres.author = 'Kylian Grafeille Clément';
  const W = 13.333, H = 7.5, M = 0.6;

  // ---------- helpers ----------
  const text = (s, t, o) => s.addText(t, { isTextBox: true, fontFace: SANS, color: C.text, margin: 0, valign: 'top', ...o });
  const card = (s, x, y, w, h, fill = C.surface) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: C.border, width: 1 } });
  const iconBox = (s, img, x, y, size = 0.6, bg = '17301F') => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: size, h: size, rectRadius: 0.08, fill: { color: bg }, line: { color: '2A5A36', width: 1 } });
    s.addImage({ data: img, x: x + size * 0.2, y: y + size * 0.2, w: size * 0.6, h: size * 0.6 });
  };
  const todo = (s, t, x, y, w, h = 0.35, size = 13) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.05, fill: { color: C.todoBg }, line: { color: C.todoBg } });
    text(s, t, { x: x + 0.1, y, w: w - 0.2, h, fontSize: size, italic: true, color: C.todo, valign: 'middle' });
  };
  let page = 0;
  const newSlide = (num, title) => {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    page++;
    if (title) {
      text(s, '// ' + num, { x: M, y: 0.45, w: 3, h: 0.35, fontFace: MONO, fontSize: 14, color: C.green });
      text(s, title, { x: M, y: 0.78, w: W - 2 * M, h: 0.75, fontSize: 34, bold: true, color: C.white });
      text(s, String(page), { x: W - M - 1, y: H - 0.45, w: 1, h: 0.3, fontFace: MONO, fontSize: 10, color: C.muted, align: 'right' });
    }
    return s;
  };
  const chip = (s, t, x, y, w) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.36, rectRadius: 0.06, fill: { color: '13233A' }, line: { color: '24466E', width: 1 } });
    text(s, t, { x, y, w, h: 0.36, fontFace: MONO, fontSize: 11, color: 'A5D6FF', align: 'center', valign: 'middle' });
  };

  // ================= 1. TITRE =================
  {
    const s = newSlide();
    text(s, '$ whoami', { x: M + 0.2, y: 1.0, w: 5, h: 0.4, fontFace: MONO, fontSize: 18, color: C.green });
    text(s, 'Kylian Grafeille\nClément', { x: M + 0.2, y: 1.45, w: 7.6, h: 1.6, fontSize: 44, bold: true, color: C.white });
    text(s, 'Alternant Support IT chez SIKA France', { x: M + 0.2, y: 3.2, w: 7.5, h: 0.45, fontSize: 20, color: C.muted });
    // fenêtre terminal
    const tx = M + 0.2, ty = 3.95, tw = 6.2;
    card(s, tx, ty, tw, 2.0);
    s.addShape(pres.shapes.RECTANGLE, { x: tx + 0.02, y: ty + 0.02, w: tw - 0.04, h: 0.4, fill: { color: C.surface2 }, line: { color: C.surface2 } });
    ['F85149', 'D29922', '3FB950'].forEach((c, i) => s.addShape(pres.shapes.OVAL, { x: tx + 0.2 + i * 0.22, y: ty + 0.14, w: 0.13, h: 0.13, fill: { color: c }, line: { color: c } }));
    text(s, 'kylian@bts-sio:~', { x: tx + 1.0, y: ty + 0.06, w: 3, h: 0.3, fontFace: MONO, fontSize: 11, color: C.muted, valign: 'middle' });
    [['formation:   ', 'BTS SIO option SISR'], ['poste:       ', 'Support IT / Réseaux'], ['entreprise:  ', 'SIKA France'], ['localisation:', 'Paris, France']].forEach(([k, v], i) =>
      text(s, [{ text: k + ' ', options: { color: C.blue } }, { text: v, options: { color: C.text } }],
        { x: tx + 0.3, y: ty + 0.58 + i * 0.33, w: tw - 0.5, h: 0.32, fontFace: MONO, fontSize: 14 }));
    // photo
    s.addShape(pres.shapes.OVAL, { x: 8.95, y: 1.55, w: 3.5, h: 3.5, fill: { color: C.surface }, line: { color: C.green, width: 2.5 } });
    text(s, 'Photo de profil\nà ajouter', { x: 8.95, y: 2.9, w: 3.5, h: 0.8, fontSize: 14, color: C.muted, align: 'center', valign: 'middle' });
    text(s, 'BTS SIO · Option SISR · Lycée Robert Schuman', { x: 8.2, y: 5.4, w: 5, h: 0.35, fontFace: MONO, fontSize: 12, color: C.muted, align: 'center' });
    s.addNotes('Présentation : nom, formation, entreprise d\'alternance. Remplacer le cercle par la photo (Insertion > Image, puis rogner en cercle).');
  }

  // ================= 2. SOMMAIRE =================
  {
    const s = newSlide('00', 'Sommaire');
    const items = ['Présentation personnelle', 'Parcours de formation', 'Projet professionnel', "Présentation de l'entreprise", 'Compétences acquises',
      'Synthèse des compétences de la formation', 'Certifications', 'Outils', 'Veille technologique', 'Connecte'];
    const colW = (W - 2 * M - 0.3) / 2;
    items.forEach((t, i) => {
      const col = i < 5 ? 0 : 1, row = i % 5;
      const x = M + col * (colW + 0.3), y = 1.85 + row * 0.98;
      card(s, x, y, colW, 0.78);
      text(s, String(i + 1).padStart(2, '0'), { x: x + 0.3, y, w: 0.6, h: 0.78, fontFace: MONO, fontSize: 18, color: C.green, valign: 'middle' });
      text(s, t, { x: x + 1.0, y, w: colW - 1.2, h: 0.78, fontSize: 18, bold: true, valign: 'middle' });
    });
  }

  // ================= 3. PRÉSENTATION PERSONNELLE =================
  {
    const s = newSlide('01', 'Présentation personnelle');
    const y = 1.85;
    // Profil
    card(s, M, y, 4.0, 4.9);
    iconBox(s, I.user, M + 0.35, y + 0.35);
    text(s, 'Profil', { x: M + 1.15, y: y + 0.4, w: 2.6, h: 0.5, fontSize: 20, bold: true, valign: 'middle' });
    [['Nom', 'Kylian Grafeille Clément'], ['Âge', '19 ans'], ['Localisation', 'Paris, France'], ['Formation', 'BTS SIO option SISR']].forEach(([k, v], i) => {
      text(s, k, { x: M + 0.35, y: y + 1.3 + i * 0.85, w: 3.3, h: 0.3, fontFace: MONO, fontSize: 11, color: C.muted });
      text(s, v, { x: M + 0.35, y: y + 1.6 + i * 0.85, w: 3.4, h: 0.4, fontSize: 17 });
    });
    // Centres d'intérêt
    const x2 = M + 4.35;
    card(s, x2, y, 3.8, 2.3);
    text(s, "Centres d'intérêt", { x: x2 + 0.35, y: y + 0.3, w: 3.2, h: 0.45, fontSize: 20, bold: true });
    chip(s, 'Voyages', x2 + 0.35, y + 1.0, 1.35); chip(s, 'Jeux vidéo', x2 + 1.85, y + 1.0, 1.6);
    chip(s, 'Sorties entre amis', x2 + 0.35, y + 1.5, 2.4);
    // Spécialité
    card(s, x2, y + 2.6, 3.8, 2.3);
    text(s, 'Spécialité', { x: x2 + 0.35, y: y + 2.9, w: 3.2, h: 0.45, fontSize: 20, bold: true });
    text(s, [{ text: 'Support informatique', options: { bold: true, color: C.green } }, { text: '\net ' }, { text: 'réseaux', options: { bold: true, color: C.green } }],
      { x: x2 + 0.35, y: y + 3.4, w: 3.2, h: 1.3, fontSize: 18 });
    // Pourquoi l'informatique
    const x3 = x2 + 4.15, w3 = W - M - x3;
    card(s, x3, y, w3, 4.9, '13261A');
    iconBox(s, I.monitor, x3 + 0.35, y + 0.35);
    text(s, "Pourquoi l'informatique ?", { x: x3 + 0.35, y: y + 1.2, w: w3 - 0.7, h: 1.0, fontSize: 20, bold: true });
    text(s, [
      { text: "J'ai découvert le fonctionnement des ordinateurs en ", options: {} },
      { text: 'montant et démontant des tours PC', options: { bold: true, color: C.white } },
      { text: '.', options: { breakLine: true } },
      { text: ' ', options: { breakLine: true, fontSize: 8 } },
      { text: 'Des membres de ma famille qui travaillent dans le domaine m\'ont également transmis cet intérêt.' },
    ], { x: x3 + 0.35, y: y + 2.35, w: w3 - 0.7, h: 2.4, fontSize: 15, color: C.text });
  }

  // ================= 4. PARCOURS =================
  {
    const s = newSlide('02', 'Parcours de formation');
    const steps = [
      { d: '2022 – 2025', t: 'Bac Pro SN', sub: 'Systèmes Numériques\nLycée Robert Schuman', b: 'Diplôme obtenu en 2025', c: C.green, bg: '17301F' },
      { d: '2025 – 2027', t: 'BTS SIO – option SISR', sub: 'Lycée Robert Schuman\nAlternance chez SIKA France', b: 'En cours', c: C.blue, bg: '13233A' },
      { d: 'Après 2027', t: 'Licence professionnelle', sub: 'Poursuite d\'études envisagée\nÉcole et formation pas encore choisies', b: 'Envisagé', c: C.muted, bg: C.surface },
    ];
    const lineY = 2.55, cw = 3.7, gap = (W - 2 * M - 3 * cw) / 2;
    s.addShape(pres.shapes.LINE, { x: M + cw / 2, y: lineY, w: 2 * (cw + gap), h: 0, line: { color: C.border, width: 2 } });
    steps.forEach((st, i) => {
      const x = M + i * (cw + gap), cx = x + cw / 2;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.2, y: lineY - 0.2, w: 0.4, h: 0.4, fill: { color: i === 2 ? C.bg : st.c }, line: { color: st.c, width: 2.5, dashType: i === 2 ? 'dash' : 'solid' } });
      text(s, st.d, { x, y: 1.85, w: cw, h: 0.35, fontFace: MONO, fontSize: 14, color: st.c, align: 'center' });
      card(s, x, 3.1, cw, 3.4);
      text(s, st.t, { x: x + 0.3, y: 3.4, w: cw - 0.6, h: 0.8, fontSize: 20, bold: true, color: C.white });
      text(s, st.sub, { x: x + 0.3, y: 4.25, w: cw - 0.6, h: 1.1, fontSize: 14, color: C.muted });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: 5.65, w: cw - 0.6, h: 0.45, rectRadius: 0.2, fill: { color: st.bg }, line: { color: st.c, width: 1, dashType: i === 2 ? 'dash' : 'solid' } });
      text(s, st.b, { x: x + 0.3, y: 5.65, w: cw - 0.6, h: 0.45, fontFace: MONO, fontSize: 12, color: st.c, align: 'center', valign: 'middle' });
    });
  }

  // ================= 5. PROJET PROFESSIONNEL =================
  {
    const s = newSlide('03', 'Projet professionnel');
    // grande carte métier
    card(s, M, 1.85, 5.6, 4.9, '13261A');
    iconBox(s, I.target, M + 0.4, 2.2, 0.7);
    text(s, 'Métier envisagé', { x: M + 0.4, y: 3.15, w: 4.8, h: 0.4, fontFace: MONO, fontSize: 13, color: C.muted });
    text(s, 'Technicien informatique', { x: M + 0.4, y: 3.55, w: 4.9, h: 1.2, fontSize: 30, bold: true, color: C.green });
    text(s, "J'aime l'informatique, et plus particulièrement le support informatique et les réseaux.",
      { x: M + 0.4, y: 5.0, w: 4.8, h: 1.2, fontSize: 16 });
    // droite : deux blocs
    const x = M + 6.0, w = W - M - x;
    [[I.book, 'Pourquoi le BTS SIO SISR ?', "L'option SISR (Solutions d'Infrastructure, Systèmes et Réseaux) est orientée vers les systèmes, les réseaux et les infrastructures informatiques : elle correspond directement à mon projet."],
     [I.trending, 'Après le BTS', 'Je souhaite éventuellement poursuivre avec une licence professionnelle.']].forEach(([ic, t, d], i) => {
      const y = 1.85 + i * 2.55;
      card(s, x, y, w, 2.35);
      iconBox(s, ic, x + 0.35, y + 0.35);
      text(s, t, { x: x + 1.15, y: y + 0.35, w: w - 1.4, h: 0.6, fontSize: 20, bold: true, valign: 'middle' });
      text(s, d, { x: x + 0.35, y: y + 1.1, w: w - 0.7, h: 1.1, fontSize: 15 });
    });
  }

  // ================= 6. COMPÉTENCES NÉCESSAIRES =================
  {
    const s = newSlide('03', 'Compétences nécessaires au métier');
    const groups = [
      ['Réseaux', I.wifi, ['Réseaux', 'Adressage IP', 'Switchs Cisco, 3Com, Netgear', 'Routeurs D-Link, Netgear', 'Modems / ADSL', "Points d'accès Wi-Fi", 'DHCP', 'DNS']],
      ['Systèmes', I.server, ['Windows 7 / 10 / 11', 'Linux Debian', 'Windows Server 2022', 'Active Directory']],
      ['Virtualisation', I.grid, ['VMware', 'VirtualBox', 'Hyper-V']],
      ['Support & sécurité', I.shield, ['Gestion de parc', 'Messagerie', 'Pack Office', 'Cybersécurité']],
    ];
    const gw = (W - 2 * M - 3 * 0.3) / 4;
    groups.forEach(([t, ic, list], i) => {
      const x = M + i * (gw + 0.3);
      card(s, x, 1.85, gw, 5.1);
      iconBox(s, ic, x + 0.3, 2.15);
      text(s, t, { x: x + 0.3, y: 2.95, w: gw - 0.6, h: 0.45, fontSize: 16, bold: true });
      list.forEach((c, j) => chip(s, c, x + 0.3, 3.6 + j * 0.41, gw - 0.6));
    });
  }

  // ================= 7. ENTREPRISE =================
  {
    const s = newSlide('04', "Présentation de l'entreprise");
    // logo placeholder
    card(s, M, 1.85, 3.2, 2.4);
    text(s, 'Logo SIKA France\nà ajouter', { x: M, y: 1.85, w: 3.2, h: 2.4, fontSize: 14, color: C.muted, align: 'center', valign: 'middle' });
    text(s, 'SIKA FRANCE', { x: M + 3.6, y: 1.85, w: 8, h: 0.6, fontSize: 30, bold: true, color: C.white });
    const stats = [['Secteur', 'BTP / industrie chimique'], ['Localité', 'Le Bourget'], ['Effectif', '≈ 1 200 – 1 500']];
    const sw = (W - M - (M + 3.6) - 0.5) / 3;
    stats.forEach(([k, v], i) => {
      const x = M + 3.6 + i * (sw + 0.25);
      card(s, x, 2.65, sw, 1.6);
      text(s, k, { x: x + 0.25, y: 2.85, w: sw - 0.5, h: 0.3, fontFace: MONO, fontSize: 12, color: C.muted });
      text(s, v, { x: x + 0.25, y: 3.25, w: sw - 0.5, h: 0.8, fontSize: i === 2 ? 17 : 18, bold: true, color: i === 2 ? C.green : C.white });
    });
    if (true) text(s, 'collaborateurs', { x: M + 3.6 + 2 * (sw + 0.25) + 0.25, y: 3.8, w: sw - 0.5, h: 0.3, fontSize: 12, color: C.muted });
    // activité & organisation
    const bw = (W - 2 * M - 0.3) / 2;
    card(s, M, 4.55, bw, 2.25);
    iconBox(s, I.briefcase, M + 0.3, 4.8, 0.55);
    text(s, 'Activité principale', { x: M + 1.05, y: 4.8, w: bw - 1.3, h: 0.55, fontSize: 18, bold: true, valign: 'middle' });
    text(s, 'Fabrication et développement de produits de spécialités chimiques pour le secteur du bâtiment.', { x: M + 0.3, y: 5.5, w: bw - 0.6, h: 1.1, fontSize: 15 });
    const x2 = M + bw + 0.3;
    card(s, x2, 4.55, bw, 2.25);
    iconBox(s, I.users, x2 + 0.3, 4.8, 0.55);
    text(s, 'Organisation', { x: x2 + 1.05, y: 4.8, w: bw - 1.3, h: 0.55, fontSize: 18, bold: true, valign: 'middle' });
    text(s, 'Filiale française du groupe suisse Sika : une direction et différents services, dont le service informatique.', { x: x2 + 0.3, y: 5.45, w: bw - 0.6, h: 0.8, fontSize: 15 });
    todo(s, 'À compléter : principaux services / sites', x2 + 0.3, 6.28, 4.4, 0.34, 12);
  }

  // ================= 8. ÉQUIPE =================
  {
    const s = newSlide('04', 'Mon équipe et ma place');
    const boxes = [['Direction', 'Cyril Jaouen'], ['Service informatique', ''], ['Responsable / tuteur', 'Sylvain Negri'], ['Alternant Support IT', 'Kylian Grafeille']];
    const bx = M + 0.3, bw = 4.6, bh = 0.95, gap = 0.35;
    boxes.forEach(([r, n], i) => {
      const y = 1.85 + i * (bh + gap), me = i === 3;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y, w: bw, h: bh, rectRadius: 0.08, fill: { color: me ? '17301F' : C.surface }, line: { color: me ? C.green : C.border, width: me ? 2 : 1 } });
      text(s, r, { x: bx, y: y + (n ? 0.14 : 0), w: bw, h: n ? 0.32 : bh, fontFace: MONO, fontSize: 12, color: me ? C.green : C.muted, align: 'center', valign: 'middle' });
      if (n) text(s, n, { x: bx, y: y + 0.45, w: bw, h: 0.4, fontSize: 18, bold: true, align: 'center', color: C.white });
      if (i < 3) s.addShape(pres.shapes.LINE, { x: bx + bw / 2, y: y + bh, w: 0, h: gap, line: { color: C.green, width: 1.5, endArrowType: 'triangle' } });
    });
    const x = M + 5.6, w = W - M - x;
    card(s, x, 1.85, w, 4.9, '13261A');
    iconBox(s, I.map, x + 0.4, 2.2, 0.7);
    text(s, "Ma place dans l'entreprise", { x: x + 0.4, y: 3.15, w: w - 0.8, h: 0.55, fontSize: 22, bold: true });
    text(s, [
      { text: 'Je suis intégré au ', options: {} }, { text: 'service informatique', options: { bold: true, color: C.green } }, { text: '.', options: { breakLine: true } },
      { text: ' ', options: { breakLine: true, fontSize: 10 } },
      { text: "J'interviens également auprès des ", options: {} }, { text: "autres services de l'entreprise", options: { bold: true, color: C.green } },
      { text: ' pour leur apporter une assistance informatique.' },
    ], { x: x + 0.4, y: 3.9, w: w - 0.8, h: 2.4, fontSize: 18 });
  }

  // ================= 9. COMPÉTENCES ACQUISES =================
  {
    const s = newSlide('05', 'Compétences acquises');
    text(s, 'Grâce à ma formation et à mon alternance', { x: M, y: 1.5, w: 8, h: 0.35, fontSize: 15, color: C.muted });
    const cols = [
      [I.wifi, 'Réseaux', ['Adressage IP', 'Configuration réseau', 'Diagnostic réseau', 'Wi-Fi']],
      [I.server, 'Administration système', ['Windows', 'Linux', 'Administration des postes', 'Machines virtuelles']],
      [I.headphones, 'Support informatique', ['Assistance aux utilisateurs', 'Diagnostic', 'Résolution de problèmes', 'Installation / configuration de logiciels', 'Maintenance des postes']],
    ];
    const cw = (W - 2 * M - 0.6) / 3;
    cols.forEach(([ic, t, list], i) => {
      const x = M + i * (cw + 0.3);
      card(s, x, 2.1, cw, 3.45);
      iconBox(s, ic, x + 0.3, 2.35, 0.55);
      text(s, t, { x: x + 1.0, y: 2.35, w: cw - 1.2, h: 0.55, fontSize: 18, bold: true, valign: 'middle' });
      text(s, list.map((l, j) => ({ text: l, options: { bullet: { code: '25B8' }, breakLine: j < list.length - 1 } })),
        { x: x + 0.3, y: 3.1, w: cw - 0.5, h: 2.3, fontSize: 15, paraSpaceAfter: 4 });
    });
    text(s, 'Compétences professionnelles', { x: M, y: 5.8, w: 5, h: 0.35, fontFace: MONO, fontSize: 13, color: C.green });
    const soft = ['Travail en équipe', 'Autonomie', 'Communication', 'Organisation'];
    const sw = (W - 2 * M - 0.9) / 4;
    soft.forEach((t, i) => {
      const x = M + i * (sw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 6.25, w: sw, h: 0.6, rectRadius: 0.08, fill: { color: '13233A' }, line: { color: '24466E', width: 1 } });
      text(s, t, { x, y: 6.25, w: sw, h: 0.6, fontSize: 16, bold: true, align: 'center', valign: 'middle', color: C.white });
    });
  }

  // ================= 10. SYNTHÈSE =================
  {
    const s = newSlide('06', 'Synthèse des compétences de la formation');
    text(s, 'Bilan des deux années de BTS SIO option SISR', { x: M, y: 1.5, w: 8, h: 0.35, fontSize: 15, color: C.muted });
    const tiles = [
      [I.globe, 'Réseaux', 'Adressage IP, configuration et diagnostic réseau, Wi-Fi'],
      [I.monitor, 'Systèmes', 'Windows, Windows Server, Linux Debian'],
      [I.headphones, 'Support informatique', "Assistance, diagnostic, résolution d'incidents"],
      [I.user, 'Administration', 'Active Directory, DHCP, DNS, gestion de parc'],
      [I.shield, 'Cybersécurité', 'Bonnes pratiques, sécurisation des postes et du réseau'],
      [I.grid, 'Virtualisation', 'VMware, VirtualBox, Hyper-V'],
      [I.users, 'Travail en équipe', 'Avec le service informatique et les autres services'],
      [I.check, 'Autonomie', 'Prise en charge et suivi des demandes'],
    ];
    const tw = (W - 2 * M - 0.9) / 4, th = 2.3;
    tiles.forEach(([ic, t, d], i) => {
      const x = M + (i % 4) * (tw + 0.3), y = 2.1 + Math.floor(i / 4) * (th + 0.3);
      card(s, x, y, tw, th);
      iconBox(s, ic, x + 0.3, y + 0.3, 0.55);
      text(s, t, { x: x + 0.3, y: y + 1.0, w: tw - 0.5, h: 0.4, fontSize: 14, bold: true });
      text(s, d, { x: x + 0.3, y: y + 1.6, w: tw - 0.5, h: 0.65, fontSize: 12.5, color: C.muted });
    });
  }

  // ================= 11. CERTIFICATIONS =================
  {
    const s = newSlide('07', 'Certifications');
    const x = 3.2, w = W - 2 * 3.2, y = 2.3;
    card(s, x, y, w, 3.6);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.5, y: y + 0.6, w: 2.4, h: 2.4, rectRadius: 0.15, fill: { color: '1BA0D7' }, line: { color: '1BA0D7' } });
    s.addImage({ data: I.cisco, x: x + 0.8, y: y + 0.9, w: 1.8, h: 1.8 });
    text(s, 'Certification', { x: x + 3.3, y: y + 0.6, w: 3, h: 0.35, fontFace: MONO, fontSize: 13, color: C.muted });
    text(s, 'CCNA', { x: x + 3.3, y: y + 0.95, w: 3.5, h: 0.9, fontSize: 48, bold: true, color: C.white });
    text(s, 'Cisco Certified Network Associate', { x: x + 3.3, y: y + 1.85, w: w - 3.6, h: 0.45, fontSize: 18, color: C.muted });
    todo(s, 'Statut / date à préciser', x + 3.3, y + 2.55, 3.3, 0.42, 14);
    s.addNotes('Préciser si la certification CCNA est obtenue (avec la date) ou s\'il s\'agit des modules Cisco NetAcad suivis en formation.');
  }

  // ================= 12. OUTILS =================
  {
    const s = newSlide('08', 'Outils');
    const tools = [
      ['kalilinux', '557C94', 'Kali Linux', 'Distribution Linux dédiée aux tests de sécurité'],
      ['windows', '0078D4', 'Active Directory', 'Gestion des utilisateurs, groupes et postes du domaine'],
      ['cisco', '1BA0D7', 'Cisco IOS', 'Système des switchs et routeurs Cisco (configuration en ligne de commande)'],
      ['vmware', '607078', 'VMware', 'Virtualisation de serveurs et de postes'],
    ];
    const tw = (W - 2 * M - 0.9) / 4;
    tools.forEach(([ic, col, t, d], i) => {
      const x = M + i * (tw + 0.3), y = 2.0;
      card(s, x, y, tw, 4.6);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + tw / 2 - 0.75, y: y + 0.55, w: 1.5, h: 1.5, rectRadius: 0.15, fill: { color: col }, line: { color: col } });
      s.addImage({ data: I[ic], x: x + tw / 2 - 0.47, y: y + 0.83, w: 0.94, h: 0.94 });
      text(s, t, { x: x + 0.2, y: y + 2.4, w: tw - 0.4, h: 0.5, fontSize: 18, bold: true, align: 'center', color: C.white });
      text(s, d, { x: x + 0.3, y: y + 3.05, w: tw - 0.6, h: 1.3, fontSize: 14, color: C.muted, align: 'center' });
    });
  }

  // ================= 13. VEILLE 1 =================
  {
    const s = newSlide('09', 'Veille technologique');
    card(s, M, 1.8, W - 2 * M, 1.0, '13233A');
    text(s, 'Sujet', { x: M + 0.35, y: 1.9, w: 3, h: 0.3, fontFace: MONO, fontSize: 12, color: C.blue });
    text(s, 'La cybersécurité des objets connectés', { x: M + 0.35, y: 2.2, w: 11, h: 0.5, fontSize: 24, bold: true, color: C.white });
    const cw = (W - 2 * M - 0.6) / 3, y = 3.1, h = 3.7;
    // pourquoi
    card(s, M, y, cw, h);
    iconBox(s, I.alert, M + 0.3, y + 0.3, 0.55);
    text(s, 'Pourquoi ce sujet ?', { x: M + 1.0, y: y + 0.3, w: cw - 1.2, h: 0.55, fontSize: 17, bold: true, valign: 'middle' });
    text(s, 'Caméras, imprimantes, badgeuses, capteurs… de plus en plus présents et souvent mal sécurisés : une porte d\'entrée sur le réseau.\n\nLien direct avec le BTS SIO SISR et mon projet de technicien support et réseaux.',
      { x: M + 0.3, y: y + 1.05, w: cw - 0.6, h: 2.5, fontSize: 13.5 });
    // Feedly
    const x2 = M + cw + 0.3;
    card(s, x2, y, cw, h);
    iconBox(s, I.rss, x2 + 0.3, y + 0.3, 0.55);
    text(s, 'Outil : Feedly', { x: x2 + 1.0, y: y + 0.3, w: cw - 1.2, h: 0.55, fontSize: 17, bold: true, valign: 'middle' });
    text(s, 'Agrégateur de flux RSS : il regroupe au même endroit les articles de mes sources pour suivre l\'actualité du sujet et conserver les articles importants.',
      { x: x2 + 0.3, y: y + 1.05, w: cw - 0.6, h: 1.6, fontSize: 13.5 });
    todo(s, 'Capture Feedly à ajouter', x2 + 0.3, y + 2.95, cw - 0.6, 0.4, 12);
    // Sources
    const x3 = x2 + cw + 0.3;
    card(s, x3, y, cw, h);
    iconBox(s, I.book, x3 + 0.3, y + 0.3, 0.55);
    text(s, 'Sources suivies', { x: x3 + 1.0, y: y + 0.3, w: cw - 1.2, h: 0.55, fontSize: 17, bold: true, valign: 'middle' });
    ['CERT-FR', 'RSS – Actualité CNIL', 'IoT Security Foundation', 'OWASP Gen AI Security Project'].forEach((t, i) =>
      chip(s, t, x3 + 0.3, y + 1.1 + i * 0.55, cw - 0.6));
  }

  // ================= 14. VEILLE 2 =================
  {
    const s = newSlide('09', 'Veille : présentation du sujet');
    const news = [
      [I.camera, 'Article marquant', 'Le botnet Mirai (2016)', 'Des milliers de caméras et box utilisant des mots de passe par défaut piratées pour mener une attaque DDoS massive.'],
      [I.zap, 'Dernière avancée', 'Directive RED', "Depuis le 1er août 2025, les équipements radio vendus dans l'UE doivent respecter des exigences de cybersécurité."],
      [I.file, 'Nouvelle importante', 'Cyber Resilience Act', 'Règlement européen en vigueur depuis décembre 2024 : produits connectés sécurisés dès la conception et mis à jour.'],
      [I.alert, 'Actualité récente', 'CRA : signalement des failles', 'Depuis septembre 2026, les fabricants doivent signaler les vulnérabilités activement exploitées.'],
    ];
    const cw = (W - 2 * M - 0.3) / 2, ch = 2.35;
    news.forEach(([ic, tag, t, d], i) => {
      const x = M + (i % 2) * (cw + 0.3), y = 1.85 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch);
      iconBox(s, ic, x + 0.3, y + 0.3, 0.7);
      text(s, tag, { x: x + 1.25, y: y + 0.3, w: cw - 1.5, h: 0.3, fontFace: MONO, fontSize: 12, color: C.green });
      text(s, t, { x: x + 1.25, y: y + 0.6, w: cw - 1.5, h: 0.45, fontSize: 19, bold: true, color: C.white });
      text(s, d, { x: x + 1.25, y: y + 1.15, w: cw - 1.55, h: 1.05, fontSize: 14, color: C.muted });
    });
    s.addNotes('Remplacer ou compléter ces exemples par les articles réellement lus sur Feedly, avec leur source.');
  }

  // ================= 15. CONNECTE =================
  {
    const s = newSlide('10', 'Connecte');
    text(s, 'Merci de votre attention', { x: M, y: 1.95, w: W - 2 * M, h: 0.7, fontSize: 30, bold: true, color: C.green });
    text(s, 'Retrouvez-moi sur :', { x: M, y: 2.7, w: 6, h: 0.4, fontSize: 16, color: C.muted });
    const links = [
      ['github', '6E5494', 'GitHub', 'github.com/Kaizoxxx', 'https://github.com/Kaizoxxx'],
      ['linkedin', '0A66C2', 'LinkedIn', 'linkedin.com/in/kylian-grafeille-clement', 'https://www.linkedin.com/in/kylian-grafeille-clement-a70776385/'],
    ];
    const cw = (W - 2 * M - 0.3) / 2;
    links.forEach(([ic, col, t, u, href], i) => {
      const x = M + i * (cw + 0.3), y = 3.5;
      card(s, x, y, cw, 2.2);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.45, y: y + 0.45, w: 1.3, h: 1.3, rectRadius: 0.15, fill: { color: col }, line: { color: col } });
      s.addImage({ data: I[ic], x: x + 0.72, y: y + 0.72, w: 0.76, h: 0.76, hyperlink: { url: href } });
      text(s, t, { x: x + 2.1, y: y + 0.55, w: cw - 2.4, h: 0.55, fontSize: 24, bold: true, color: C.white });
      text(s, [{ text: u, options: { hyperlink: { url: href }, color: C.blue } }], { x: x + 2.1, y: y + 1.15, w: cw - 2.3, h: 0.45, fontFace: MONO, fontSize: 13 });
    });
  }

  await pres.writeFile({ fileName: OUT });
  console.log('écrit :', OUT);
})();
