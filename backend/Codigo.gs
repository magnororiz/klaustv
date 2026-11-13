/**
 * Klaus TV — backend de login (Google Apps Script)
 * Planilha: "Klaus TV - Banco do Site", aba "Usuarios".
 * Colunas: A Usuário | B Nome | C Perfil (Dono/Operador) | D Senha (digite aqui; o sistema protege e apaga)
 *          E Ativo (SIM/NÃO) | F Páginas liberadas | G E-mail | H Senha protegida | I Trocar senha no próximo acesso | J Último acesso
 * Publicar: Implantar > Nova implantação > App da Web > Executar como: Eu > Quem pode acessar: Qualquer pessoa.
 */
const PLANILHA_ID = '17x42osZNhtjsPWIKIA5NHtjr6gs6Iiz7nw4I_YXvNFQ';
const ABA = 'Usuarios';
const SITE = 'https://klaustv.com';
const HORAS_SESSAO = 12;

function doGet(e) { return responder({ ok: true, msg: 'Klaus TV backend ativo' }); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents || '{}'); } catch (err) { return responder({ ok: false, erro: 'Pedido inválido.' }); }
  if (p.acao === 'aoVivoLista') return responder(aoVivoLista(p.categoria));
  if (p.acao === 'aoVivoJogo') return responder(aoVivoJogo(p.id));
  if (p.acao === 'dadosFed') return responder(dadosFed());
  if (p.acao === 'dadosCestinhas') return responder(dadosCestinhas());
  if (p.acao === 'dadosSumulas') return responder(dadosSumulas());
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    switch (p.acao) {
      case 'login': return responder(login(p.usuario, p.senha));
      case 'trocarSenha': return responder(trocarSenha(p.token, p.senhaAtual, p.senhaNova));
      case 'eu': return responder(eu(p.token));
      case 'listarUsuarios': return responder(listarUsuarios(p.token));
      case 'criarUsuario': return responder(criarUsuario(p.token, p));
      case 'excluirUsuario': return responder(excluirUsuario(p.token, p.usuario));
      case 'salvarJogo': return responder(salvarJogo(p.token, p.jogo));
      case 'listarJogos': return responder(listarJogos(p.token, p.data));
      case 'excluirJogo': return responder(excluirJogo(p.token, p.id));
      case 'salvarStats': return responder(salvarStats(p.token, p.id, p.stats, p.status, p.dataHora));
      case 'salvarFoto': return responder(salvarFoto(p.token, p.chave, p.imagem));
      case 'listarFotos': return responder(listarFotos(p.token));
      case 'salvarApelido': return responder(salvarApelido(p.token, p.reg, p.apelido));
      case 'salvarCamisa': return responder(salvarCamisa(p.token, p.reg, p.camisa));
      case 'aoVivoLista': return responder(aoVivoLista(p.categoria));
      case 'aoVivoJogo': return responder(aoVivoJogo(p.id));
      case 'listarArbitragem': return responder(listarArbitragem(p.token));
      case 'incluirArbitragem': return responder(incluirArbitragem(p.token, p.nomes));
      case 'funcaoArbitragem': return responder(funcaoArbitragem(p.token, p.nome, p.funcao));
      case 'excluirArbitragem': return responder(excluirArbitragem(p.token, p.nome));
      case 'statusBackup': return responder(statusBackup(p.token));
      case 'rodarBackup': return responder(rodarBackupSite(p.token));
      case 'sair': CacheService.getScriptCache().remove('s_' + p.token); return responder({ ok: true });
      default: return responder({ ok: false, erro: 'Ação desconhecida.' });
    }
  } catch (err) {
    return responder({ ok: false, erro: 'Erro no servidor: ' + err.message });
  } finally { lock.releaseLock(); }
}

function responder(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- planilha ---------- */
function aba() { return SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA); }
function linhas() {
  const sh = aba(); const v = sh.getDataRange().getValues();
  return v.slice(1).map((r, i) => ({
    linha: i + 2, usuario: String(r[0]).trim().toLowerCase(), nome: r[1], perfil: String(r[2]).trim(),
    senha: String(r[3]), ativo: String(r[4]).trim().toUpperCase() !== 'NÃO', paginas: r[5], email: String(r[6]).trim(),
    hash: String(r[7]), trocar: String(r[8]).trim().toUpperCase() === 'SIM'
  })).filter(u => u.usuario);
}
function achar(usuario) { usuario = String(usuario || '').trim().toLowerCase(); return linhas().find(u => u.usuario === usuario); }

/* ---------- senha ---------- */
function gerarSal() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function hashSenha(senha, sal) {
  let d = sal + '|' + senha;
  for (let i = 0; i < 1000; i++) {
    d = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d, Utilities.Charset.UTF_8));
  }
  return sal + '$' + d;
}
function confere(senha, guardado) { const sal = String(guardado).split('$')[0]; return hashSenha(senha, sal) === guardado; }
function senhaProvisoria() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789'; let s = '';
  for (let i = 0; i < 8; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return s;
}
function senhaValida(s) { return typeof s === 'string' && s.length >= 6; }

/* Protege as senhas digitadas na coluna D (pode rodar manualmente também). */
function protegerSenhas() {
  const sh = aba();
  linhas().forEach(u => {
    if (u.senha && u.senha !== '') {
      sh.getRange(u.linha, 8).setValue(hashSenha(u.senha, gerarSal()));
      sh.getRange(u.linha, 4).setValue('');
    }
  });
}

/* ---------- sessão ---------- */
function novaSessao(u) {
  const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  CacheService.getScriptCache().put('s_' + token, u.usuario, HORAS_SESSAO * 3600);
  return token;
}
function daSessao(token) {
  if (!token) return null;
  const usuario = CacheService.getScriptCache().get('s_' + token);
  if (!usuario) return null;
  const u = achar(usuario); return (u && u.ativo) ? u : null;
}
function publico(u) { return { usuario: u.usuario, nome: u.nome, perfil: u.perfil, email: u.email, trocarSenha: u.trocar, podeGerenciar: u.perfil === 'Dono' }; }

/* ---------- ações ---------- */
function login(usuario, senha) {
  protegerSenhas();
  const u = achar(usuario);
  if (!u || !u.ativo || !u.hash || !confere(String(senha || ''), u.hash)) return { ok: false, erro: 'Usuário ou senha incorretos.' };
  aba().getRange(u.linha, 10).setValue(new Date());
  return { ok: true, token: novaSessao(u), usuario: publico(u) };
}

function eu(token) { const u = daSessao(token); return u ? { ok: true, usuario: publico(u) } : { ok: false, erro: 'Sessão expirada. Entre de novo.' }; }

function trocarSenha(token, atual, nova) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  if (!confere(String(atual || ''), u.hash)) return { ok: false, erro: 'A senha atual está errada.' };
  if (!senhaValida(nova)) return { ok: false, erro: 'A nova senha precisa ter pelo menos 6 caracteres.' };
  if (nova === atual) return { ok: false, erro: 'A nova senha precisa ser diferente da atual.' };
  const sh = aba();
  sh.getRange(u.linha, 8).setValue(hashSenha(nova, gerarSal()));
  sh.getRange(u.linha, 9).setValue('NÃO');
  return { ok: true, usuario: publico(achar(u.usuario)) };
}

function exigeDono(token) {
  const u = daSessao(token);
  if (!u) return { erro: { ok: false, erro: 'Sessão expirada. Entre de novo.' } };
  if (u.perfil !== 'Dono') return { erro: { ok: false, erro: 'Só o Klaus pode gerenciar usuários.' } };
  return { u: u };
}

function listarUsuarios(token) {
  const d = exigeDono(token); if (d.erro) return d.erro;
  return { ok: true, usuarios: linhas().map(u => ({ usuario: u.usuario, nome: u.nome, perfil: u.perfil, email: u.email, ativo: u.ativo, trocarSenha: u.trocar })) };
}

function criarUsuario(token, p) {
  const d = exigeDono(token); if (d.erro) return d.erro;
  const usuario = String(p.usuario || '').trim().toLowerCase();
  const nome = String(p.nome || '').trim() || usuario;
  const senhaDada = String(p.senha || '');
  const email = String(p.email || '').trim();
  if (!/^[a-z0-9._-]{3,30}$/.test(usuario)) return { ok: false, erro: 'Login inválido: use de 3 a 30 letras minúsculas, números, ponto ou traço.' };
  if (senhaDada && !senhaValida(senhaDada)) return { ok: false, erro: 'A senha provisória precisa ter pelo menos 6 caracteres.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, erro: 'E-mail inválido.' };
  if (achar(usuario)) return { ok: false, erro: 'Já existe um usuário com esse login.' };
  const prov = senhaDada || senhaProvisoria();
  aba().appendRow([usuario, nome, 'Operador', '', 'SIM', 'Scout, Ao vivo, Estatísticas', email, hashSenha(prov, gerarSal()), 'SIM', '']);
  MailApp.sendEmail({
    to: email,
    subject: 'Klaus TV — seu acesso ao site',
    htmlBody: '<p>Olá, ' + nome + '!</p><p>O Klaus criou seu acesso ao site <b>Klaus TV</b>.</p>' +
      '<p>Endereço: <a href="' + SITE + '">' + SITE + '</a><br>Login: <b>' + usuario + '</b><br>Senha provisória: <b>' + prov + '</b></p>' +
      '<p>No primeiro acesso o site vai pedir para você criar uma senha nova.</p><p>Klaus TV · Basquete do Rio de Janeiro</p>',
    name: 'Klaus TV'
  });
  return { ok: true, msg: 'Usuário criado e e-mail enviado para ' + email + '.' };
}

function excluirUsuario(token, usuario) {
  const d = exigeDono(token); if (d.erro) return d.erro;
  const u = achar(usuario); if (!u) return { ok: false, erro: 'Usuário não encontrado.' };
  if (u.perfil === 'Dono') return { ok: false, erro: 'O usuário do Klaus não pode ser excluído.' };
  aba().deleteRow(u.linha);
  return { ok: true, msg: 'Usuário ' + u.usuario + ' excluído.' };
}

/* Rode uma vez pelo editor para autorizar planilha e e-mail. */
function autorizar() { protegerSenhas(); pastaFotos(); Logger.log(MailApp.getRemainingDailyQuota() + ' e-mails disponíveis hoje'); }


/* ================= JOGOS (Montar Jogo / Scout / Ao vivo) ================= */
const ABA_JOGOS = 'JogosKT';
const COLS = ['id','categoria','dataHora','timeA','timeB','ginasio','jogoFed','meta','stats','status','placarA','placarB','atualizado','criadoPor'];
function abaJogos() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(ABA_JOGOS);
  if (!sh) { sh = ss.insertSheet(ABA_JOGOS); sh.getRange(1,1,1,COLS.length).setValues([COLS]); sh.setFrozenRows(1); }
  return sh;
}
function lerJogos() {
  const v = abaJogos().getDataRange().getValues();
  return v.slice(1).map((r, i) => { const o = { linha: i + 2 }; COLS.forEach((c, j) => o[c] = r[j]); return o; }).filter(o => o.id);
}
function jsonOu(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }
function placarDe(stats) {
  const t = (lado) => { let p = 0; const o = (stats && stats[lado]) || {}; Object.keys(o).forEach(k => p += Number(o[k].pts || 0)); return p; };
  return [t('A'), t('B')];
}
function salvarJogo(token, j) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  if (!j || !j.timeA || !j.timeB || !j.categoria) return { ok: false, erro: 'Jogo incompleto.' };
  const sh = abaJogos(); const todos = lerJogos();
  const id = j.id || ('KT-' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 1000));
  const atual = todos.find(x => x.id === id);
  const meta = { timeA: j.timeA, timeB: j.timeB, oficiais: j.oficiais || {}, ginasio: j.ginasio || '', mandante: j.mandante || '', jogoFed: j.jogoFed || '', dataFed: j.dataFed || '', avulso: !!j.avulso, atualizadoEm: new Date().toISOString() };
  const linha = [id, j.categoria, j.dataHora || '', j.timeA.sigla || '', j.timeB.sigla || '', j.ginasio || '', j.jogoFed || '', JSON.stringify(meta),
    atual ? atual.stats : '', atual ? atual.status : '', atual ? atual.placarA : 0, atual ? atual.placarB : 0, new Date(), u.usuario];
  if (atual) sh.getRange(atual.linha, 1, 1, COLS.length).setValues([linha]); else sh.appendRow(linha);
  if (j.avulso) gravarAvulso(id, j, meta, u.usuario);
  CacheService.getScriptCache().remove('j_' + id); CacheService.getScriptCache().remove('lista');
  return { ok: true, id: id };
}
function listarJogos(token, data) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  let js = lerJogos();
  if (data) js = js.filter(x => String(x.dataHora).slice(0, 10) === data);
  return { ok: true, jogos: js.map(x => ({ id: x.id, categoria: x.categoria, dataHora: x.dataHora, ginasio: x.ginasio, jogoFed: x.jogoFed, meta: jsonOu(x.meta, {}), status: x.status, placarA: x.placarA, placarB: x.placarB })) };
}
function excluirJogo(token, id) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  const j = lerJogos().find(x => x.id === id); if (!j) return { ok: false, erro: 'Jogo não encontrado.' };
  abaJogos().deleteRow(j.linha); apagarAvulso(id); CacheService.getScriptCache().remove('lista');
  return { ok: true };
}
function salvarStats(token, id, stats, status, dataHora) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  const j = lerJogos().find(x => x.id === id); if (!j) return { ok: false, erro: 'Jogo não encontrado.' };
  const pl = placarDe(stats); const sh = abaJogos();
  sh.getRange(j.linha, 9, 1, 5).setValues([[JSON.stringify(stats || {}), status || j.status || '', pl[0], pl[1], new Date()]]);
  if (dataHora && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(String(dataHora))) { const c = sh.getRange(j.linha, 3); c.setNumberFormat('@'); c.setValue(String(dataHora).slice(0, 16)); }
  const c = CacheService.getScriptCache(); c.remove('j_' + id); c.remove('lista');
  return { ok: true, placar: pl };
}
function aoVivoLista(categoria) {
  const c = CacheService.getScriptCache(); let lista = jsonOu(c.get('lista'), null);
  if (!lista) {
    lista = lerJogos().map(x => ({ id: x.id, categoria: x.categoria, dataHora: x.dataHora instanceof Date ? x.dataHora.toISOString() : String(x.dataHora),
      timeA: x.timeA, timeB: x.timeB, placarA: Number(x.placarA || 0), placarB: Number(x.placarB || 0), status: x.status, ginasio: x.ginasio, avulso: !!jsonOu(x.meta, {}).avulso }));
    c.put('lista', JSON.stringify(lista), 20);
  }
  if (categoria === '__AV') lista = lista.filter(x => x.avulso);
  else if (categoria) lista = lista.filter(x => x.categoria === categoria && !x.avulso);
  return { ok: true, jogos: lista, agora: new Date().toISOString() };
}
function aoVivoJogo(id) {
  const c = CacheService.getScriptCache(); const k = 'j_' + id; const hit = c.get(k);
  if (hit) return JSON.parse(hit);
  const x = lerJogos().find(j => j.id === id); if (!x) return { ok: false, erro: 'Jogo não encontrado.' };
  const fotos = mapaFotos();
  const r = { ok: true, fotos: fotos, apelidos: mapaApelidos(), jogo: { id: x.id, categoria: x.categoria, dataHora: x.dataHora instanceof Date ? x.dataHora.toISOString() : String(x.dataHora),
    ginasio: x.ginasio, jogoFed: x.jogoFed, meta: jsonOu(x.meta, {}), stats: jsonOu(x.stats, {}), status: x.status }, agora: new Date().toISOString() };
  c.put(k, JSON.stringify(r), 4);
  return r;
}


/* ================= FOTOS (jogadores e técnicos) ================= */
const ABA_FOTOS = 'Fotos';
const PASTA_FOTOS = 'Klaus TV - Fotos dos atletas';
function abaFotos() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName(ABA_FOTOS);
  if (!sh) { sh = ss.insertSheet(ABA_FOTOS); sh.getRange(1,1,1,4).setValues([['chave','url','arquivoId','atualizado']]); sh.setFrozenRows(1); }
  return sh;
}
function pastaFotos() {
  const it = DriveApp.getFoldersByName(PASTA_FOTOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_FOTOS);
}
function mapaFotos() {
  const c = CacheService.getScriptCache(); const hit = c.get('fotos'); if (hit) return JSON.parse(hit);
  const v = abaFotos().getDataRange().getValues().slice(1); const m = {};
  v.forEach(r => { if (r[0]) m[String(r[0])] = r[1]; });
  c.put('fotos', JSON.stringify(m), 300);
  return m;
}
function listarFotos(token) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  return { ok: true, fotos: mapaFotos(), apelidos: mapaApelidos(), camisas: mapaCamisas() };
}
function salvarFoto(token, chave, imagem) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  chave = String(chave || '').trim(); if (!/^[A-Za-z0-9_-]{2,40}$/.test(chave)) return { ok: false, erro: 'Chave inválida.' };
  const m = String(imagem || '').match(/^data:(image\/(jpeg|png|webp));base64,(.+)$/); if (!m) return { ok: false, erro: 'Imagem inválida.' };
  const bytes = Utilities.base64Decode(m[3]); if (bytes.length > 600000) return { ok: false, erro: 'Imagem grande demais.' };
  const sh = abaFotos(); const v = sh.getDataRange().getValues();
  const i = v.findIndex((r, k) => k > 0 && String(r[0]) === chave);
  if (i > 0 && v[i][2]) { try { DriveApp.getFileById(v[i][2]).setTrashed(true); } catch (e) {} }
  const f = pastaFotos().createFile(Utilities.newBlob(bytes, m[1], chave + '.' + (m[2] === 'jpeg' ? 'jpg' : m[2])));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://lh3.googleusercontent.com/d/' + f.getId() + '=w600';
  const linha = [chave, url, f.getId(), new Date()];
  if (i > 0) sh.getRange(i + 1, 1, 1, 4).setValues([linha]); else sh.appendRow(linha);
  const c = CacheService.getScriptCache(); c.remove('fotos'); c.removeAll(['lista']);
  return { ok: true, url: url };
}


/* ================= APELIDOS (valem para todos os jogos) ================= */
function abaApelidos() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName('Apelidos');
  if (!sh) { sh = ss.insertSheet('Apelidos'); sh.getRange(1,1,1,3).setValues([['cadastro','apelido','atualizado']]); sh.setFrozenRows(1); }
  return sh;
}
function mapaApelidos() {
  const c = CacheService.getScriptCache(); const hit = c.get('apelidos'); if (hit) return JSON.parse(hit);
  const m = {}; abaApelidos().getDataRange().getValues().slice(1).forEach(r => { if (r[0] && r[1]) m[String(r[0])] = String(r[1]); });
  c.put('apelidos', JSON.stringify(m), 300); return m;
}
function salvarApelido(token, reg, apelido) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  reg = String(reg || '').trim(); apelido = String(apelido || '').trim().slice(0, 30);
  if (!/^[A-Za-z0-9_-]{2,40}$/.test(reg)) return { ok: false, erro: 'Cadastro inválido.' };
  const sh = abaApelidos(); const v = sh.getDataRange().getValues();
  const i = v.findIndex((r, k) => k > 0 && String(r[0]) === reg);
  if (i > 0) { sh.getRange(i + 1, 2, 1, 2).setValues([[apelido, new Date()]]); }
  else if (apelido) sh.appendRow([reg, apelido, new Date()]);
  CacheService.getScriptCache().remove('apelidos');
  return { ok: true };
}


/* ================= JOGOS AVULSOS (aba própria, cópia de cada jogo avulso) ================= */
function abaAvulsos() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName('JogosAvulsos');
  if (!sh) { sh = ss.insertSheet('JogosAvulsos'); sh.getRange(1,1,1,8).setValues([['id','categoria','dataHora','timeA','timeB','ginasio','dados','atualizado']]); sh.setFrozenRows(1); }
  return sh;
}
function gravarAvulso(id, j, meta, quem) {
  const sh = abaAvulsos(); const v = sh.getDataRange().getValues();
  const i = v.findIndex((r, k) => k > 0 && r[0] === id);
  const linha = [id, j.categoria, j.dataHora || '', j.timeA.nome || '', j.timeB.nome || '', j.ginasio || '', JSON.stringify(meta), new Date()];
  if (i > 0) sh.getRange(i + 1, 1, 1, 8).setValues([linha]); else sh.appendRow(linha);
}
function apagarAvulso(id) {
  const sh = abaAvulsos(); const v = sh.getDataRange().getValues();
  const i = v.findIndex((r, k) => k > 0 && r[0] === id); if (i > 0) sh.deleteRow(i + 1);
}

/* ================= CAMISA PADRÃO (Banco de jogadores) — coluna D da aba Apelidos ================= */
function mapaCamisas() {
  const c = CacheService.getScriptCache(); const hit = c.get('camisas'); if (hit) return JSON.parse(hit);
  const m = {}; abaApelidos().getDataRange().getValues().slice(1).forEach(r => { if (r[0] && r[3] !== '' && r[3] !== null && r[3] !== undefined) m[String(r[0])] = String(r[3]); });
  c.put('camisas', JSON.stringify(m), 300); return m;
}
function salvarCamisa(token, reg, camisa) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  reg = String(reg || '').trim(); camisa = String(camisa || '').replace(/[^0-9]/g, '').slice(0, 3);
  if (!/^[A-Za-z0-9_-]{2,40}$/.test(reg)) return { ok: false, erro: 'Cadastro inválido.' };
  const sh = abaApelidos(); if (sh.getRange(1, 4).getValue() !== 'camisa') sh.getRange(1, 4).setValue('camisa');
  const v = sh.getDataRange().getValues(); const i = v.findIndex((r, k) => k > 0 && String(r[0]) === reg);
  if (i > 0) { const cel = sh.getRange(i + 1, 4); cel.setNumberFormat('@'); cel.setValue(camisa); sh.getRange(i + 1, 3).setValue(new Date()); }
  else { sh.appendRow([reg, '', new Date(), camisa]); sh.getRange(sh.getLastRow(), 4).setNumberFormat('@').setValue(camisa); }
  CacheService.getScriptCache().remove('camisas');
  return { ok: true };
}


/* ---------- backup diário da FBERJ (20h) ----------
 * Elenco: planilha "Cestinhas_CEMB_2026" (xlsx) -> aba FBERJ_Atletas. SÓ ACRESCENTA, nunca apaga.
 * Tabela: planilha "CEMB 2026" -> aba FBERJ_Tabela. Só troca se a leitura nova vier completa.
 * Se a FBERJ sair do ar ou mudar o arquivo, o site continua com o último backup.
 */
const FED_CESTINHAS_ID = '1wekGITrAep1-1pAttzs0hpY9e5-pPY_8';
const FED_TABELA_ID = '1SsTnHglFy0oSvFUT6i9OH_ZBAQmxtlxVIEnFCEYKNiE';
const FED_TABELA_FEM_ID = '1AEGj7wdpXq0lRWRnv3ZMA_RVyI7qvDs2pRX08LhiiNc'; // 3ª Taça Magaly Costa (feminino)
const FED_SUMULAS_ID = '1ApNK5Ke04dY31uOZQvLfZoEFFDpmMCfy'; // pasta CEMB 2026 da FBERJ
const FED_MAP = {"ABIG - IGUABA GRANDE": "IG", "ABIG / IGUABA GRANDE": "IG", "ABIG / JIC / SHARK'S": "ABIG", "ACEAR / IRAJÁ A. C. / PJMC": "ACEAR", "ACEAR / IRAJÁ A.C. / PJMC": "ACEAR", "ACEAR / IRAJÁ A.C. / PROJETO MC": "ACEAR", "ASP C. MUNICIPAL": "ASPCM", "ASP CLUB MUNICIPAL": "ASPCM", "BASQUETE S.A": "BSA", "BASQUETE S.A.": "BSA", "BOTAFOGO F. R.": "BFR", "BOTAFOGO F.R.": "BFR", "C. R. FLAMENGO": "CRF", "C.R. FLAMENGO": "CRF", "CARIOCA KING'S": "CK", "CCL": "CCL", "CESC": "CESC", "CFCSN": "CFCSN", "E. E. PASSO ZERO": "EEPZ", "E.E. PASSO ZERO": "EEPZ", "FLUMINENSE F. C.": "FFC", "FLUMINENSE F.C.": "FFC", "FLUMINESE F. C.": "FFC", "IDEC": "IDEC", "ITAIPU BASQUETE": "IB", "ITDA": "ITDA", "JACAREPAGUÁ T. C.": "JTC", "JACAREPAGUÁ T.C.": "JTC", "K ESPORTES": "KE", "MESQUITA BASQUETE": "MB", "NBR CAMPOS": "NBR", "NITERÓI B. C.": "NBC", "NITERÓI B.C.": "NBC", "TEAM KLUN ACADEMY": "TKA", "TIJUCA BASKET": "TTC", "TIJUCA T. C. / RAIZEIRAS": "TTC"};
const FED_GYM = { 'NBA STORE ARENA': 'NBA Store Arena Rio - Uptown' };
const FED_CAT_CEST = { SUB12: 'Sub-12', SUB13: 'Sub-13', SUB14: 'Sub-14', SUB15: 'Sub-15', SUB16: 'Sub-16', SUB17: 'Sub-17', SUB19: 'Sub-19', SUB22: 'Sub-22', ADULTO: 'Adulto' };
const FED_CAT_TAB = { SUB_12: 'Sub-12', SUB_13: 'Sub-13', SUB_14: 'Sub-14', SUB_15: 'Sub-15', SUB_16: 'Sub-16', SUB_17: 'Sub-17', SUB_19: 'Sub-19', SUB_22: 'Sub-22', ADULTO: 'Adulto' };
const FED_DEF = { 'Sub-12': 'Turno', 'Sub-13': 'Turno', 'Sub-14': 'Turno', 'Sub-22': 'Turno', 'Adulto': 'Classificação' };

function fedAba(nome, cab) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID); let sh = ss.getSheetByName(nome);
  if (!sh) { sh = ss.insertSheet(nome); sh.getRange(1, 1, 1, cab.length).setValues([cab]); sh.setFrozenRows(1); }
  return sh;
}
function fedNorm(x) { return String(x == null ? '' : x).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim(); }
function fedTitulo(x) { x = fedNorm(x); if (x && x === x.toUpperCase()) x = x.toLowerCase().replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase()); return x; }

function fedConverterXlsx(id) {
  // usa o serviço avançado do Drive (já autorizado pela conta); não precisa de UrlFetch
  const c = Drive.Files.copy({ name: 'Klaus TV - backup FBERJ (temporário)', mimeType: 'application/vnd.google-apps.spreadsheet' }, id, { supportsAllDrives: true });
  if (!c || !c.id) throw new Error('Não consegui copiar a planilha da FBERJ ' + id + '.');
  return c.id;
}

function fedLerElenco() {
  const tmp = fedConverterXlsx(FED_CESTINHAS_ID); const atl = {}; const equipes = {}; const pontos = { resumo: [], jogos: [] };
  try {
    SpreadsheetApp.openById(tmp).getSheets().forEach(ws => {
      try { fedLerPontos(ws, pontos); } catch (e) {}
      const t = ws.getName().toUpperCase(); const ck = Object.keys(FED_CAT_CEST).find(k => t.indexOf(k) === 0); if (!ck) return;
      const cat = FED_CAT_CEST[ck]; const rows = ws.getDataRange().getDisplayValues();
      for (let i = 0; i < rows.length; i++) {
        if (fedNorm(rows[i][0]) !== 'Reg') continue;
        let sig = null;
        for (let k = i - 1; k >= Math.max(0, i - 4); k--) { const m = fedNorm(rows[k][0]).match(/^(.*)\(([A-Z]+)\)$/); if (m) { sig = m[2]; if (!equipes[sig]) equipes[sig] = fedNorm(m[1]); break; } }
        if (!sig) continue;
        for (let k = i + 1; k < rows.length; k++) {
          const r0 = fedNorm(rows[k][0]), nome = fedNorm(rows[k][1]);
          if (!r0 || r0.toUpperCase() === 'TOTAL' || !nome) break;
          let reg = r0.replace(/[.,]0+$/, '').split(' ')[0]; if (!/^\d+$/.test(reg)) continue;
          atl[cat + '|' + sig + '|' + reg] = { cat: cat, sig: sig, reg: reg, nome: nome };
        }
      }
    });
  } finally { try { DriveApp.getFileById(tmp).setTrashed(true); } catch (e) {} }
  return { atletas: atl, equipes: equipes, pontos: pontos };
}

/* Pontos dos cestinhas (mesma lógica do parse_fed.py):
 * resumo = total, cestas de 3, jogos e média (aba final de cada categoria)
 * jogos  = pontos de cada atleta em cada jogo: [cat, fase, ocorrencia, sigla, adversario, matricula, pts] */
const FED_FINAL = { 'Sub-12': 'SUB12P', 'Sub-13': 'SUB13P', 'Sub-14': 'SUB14P', 'Sub-15': 'SUB15S', 'Sub-16': 'SUB16S', 'Sub-17': 'SUB17S', 'Sub-19': 'SUB19S', 'Sub-22': 'SUB22P', 'Adulto': 'ADULTO' };
function fedLerPontos(ws, out) {
  const t = ws.getName().toUpperCase().trim(); const ck = Object.keys(FED_CAT_CEST).find(k => t.indexOf(k) === 0); if (!ck) return;
  const cat = FED_CAT_CEST[ck]; const rows = ws.getDataRange().getValues();
  const tit = rows[1] ? fedNorm(rows[1][0]) : ''; const mm = tit.match(/2026\s*-\s*(.*)$/);
  let def = mm ? fedTitulo(mm[1]) : 'Classificação'; if (/^TURNO\/RETURNO/i.test(def) || /^CLASSIFICA/i.test(def)) def = 'Classificação';
  const ehFinal = FED_FINAL[cat] === t;
  for (let i = 2; i < rows.length; i++) {
    if (fedNorm(rows[i][0]) !== 'Reg') continue;
    let sig = null;
    for (let k = i - 1; k >= Math.max(0, i - 4); k--) { const m = fedNorm(rows[k][0]).match(/^(.*)\(([A-Z]+)\)$/); if (m) { sig = m[2]; break; } }
    if (!sig) continue;
    const hdr = rows[i], ph = rows[i - 2], op = rows[i - 1];
    const H = hdr.map(h => typeof h === 'string' ? h.trim().toLowerCase() : '');
    const cols = []; H.forEach((h, c) => { if (h === 'pontos' || (h === 'total' && H[c + 1] === 'total' && c > 0 && H[c - 1] === 'pontos')) cols.push(c); });
    const info = {}; let lastop = null, lastph = null;
    cols.forEach(c => { let p = null; for (let cc = c; cc > 1; cc--) { if (ph[cc] !== '' && ph[cc] != null) { p = fedTitulo(String(ph[cc])); break; } } p = p || def;
      let o = (op[c] !== '' && op[c] != null) ? op[c] : (lastph === p ? lastop : null); o = o ? fedNorm(o) : '?'; info[c] = [p, o]; lastop = o; lastph = p; });
    const jcols = []; hdr.forEach((h, c) => { if (typeof h === 'string' && h.trim().indexOf('Nº de jogos') === 0) jcols.push(c); }); const j = jcols.length ? jcols[jcols.length - 1] : null;
    const tc = [], cc3 = []; if (j !== null) for (let c = 0; c < j; c++) { const h = typeof hdr[c] === 'string' ? hdr[c].trim().toLowerCase() : ''; if (h === 'total') tc.push(c); if (h.indexOf('cestas') === 0) cc3.push(c); }
    for (let k = i + 1; k < rows.length; k++) {
      const r = rows[k]; const r0 = fedNorm(r[0]); if (!r0 || r0.toUpperCase() === 'TOTAL' || r[1] === '' || r[1] == null) break;
      const reg = String(r0).replace(/[.,]0+$/, '').split(' ')[0]; const nome = fedNorm(r[1]); const cnt = {};
      cols.forEach(c => { const [p, o] = info[c]; const key = p + '|' + o; cnt[key] = (cnt[key] || 0) + 1; const v = r[c]; if (typeof v === 'number') out.jogos.push([cat, p, cnt[key], sig, o, reg, v]); });
      if (ehFinal && j !== null) { const n = x => typeof x === 'number' ? x : ''; out.resumo.push([cat, sig, reg, nome, n(r[tc[tc.length - 1]]), n(r[cc3[cc3.length - 1]]), n(r[j]), n(r[j + 1])]); }
    }
  }
}

function fedFase(cat, label) {
  let l = fedNorm(label.replace(/,/g, ' ')).toUpperCase().replace('RETIURNO', 'RETURNO').replace('SEMIS', 'SEMIFINAL').replace('SEMI ', 'SEMIFINAL ');
  if (l === 'SEMI') l = 'SEMIFINAL';
  const serie = ['Sub-15', 'Sub-16', 'Sub-17', 'Sub-19'].indexOf(cat) >= 0; const parts = l.split(' '); const head = parts[0];
  const nivel = fedTitulo(parts.slice(1).join(' '));
  if (head === 'RETURNO') return serie ? ('Série' + (nivel ? ' ' + nivel : '')) : ('Returno' + (nivel ? ' ' + nivel : ''));
  if (head.indexOf('SEMIFINAL') === 0) return 'Semifinal' + (nivel ? ' ' + nivel : '');
  if (head === 'FINAL') return 'Final' + (nivel ? ' ' + nivel : '');
  return fedTitulo(l);
}

function fedLerTabela(id, feminino) {
  const out = []; const semSigla = {}; const tmp = id ? null : fedConverterXlsx(FED_TABELA_ID);
  try {
  SpreadsheetApp.openById(tmp || id).getSheets().forEach(ws => {
    const nomeAba = ws.getName().trim().toUpperCase();
    const cat = feminino ? (nomeAba === 'ADULTO' ? 'Adulto' : (nomeAba.match(/^SUB_(\d+)$/) ? 'Sub-' + nomeAba.slice(4) : null)) : FED_CAT_TAB[nomeAba]; if (!cat) return;
    let fase = feminino ? 'Classificação' : (FED_DEF[cat] || 'Turno Único');
    ws.getDataRange().getDisplayValues().forEach(r => {
      r = r.map(fedNorm); while (r.length < 10) r.push('');
      if (!r.join('') || r[0] === 'DATA') return;
      if (!r[2] && !r[3]) { const lab = [r[0], r[1]].filter(x => x).join(' '); if (lab && lab !== '/') fase = fedFase(cat, lab); return; }
      if (r[2] === 'GINÁSIO') return;
      const a = feminino ? fedTitulo(r[4]) : (FED_MAP[r[4]] || ''), b = feminino ? fedTitulo(r[8]) : (FED_MAP[r[8]] || '');
      if (r[4] && !a) semSigla[r[4]] = 1; if (r[8] && !b) semSigla[r[8]] = 1;
      out.push([cat, fase, r[3], r[0], r[1].slice(0, 5), a, b, FED_GYM[r[2]] || r[2], r[5], r[7], r[9]]);
    });
  });
  } finally { if (tmp) try { DriveApp.getFileById(tmp).setTrashed(true); } catch (e) {} }
  return { linhas: out, semSigla: Object.keys(semSigla) };
}

function backupFBERJ() {
  const log = fedAba('FBERJ_Backup', ['quando', 'resultado', 'detalhes']); const agora = new Date(); const det = [];
  let ok = true;
  // 1) elenco: só acrescenta
  try {
    const e = fedLerElenco(); const n = Object.keys(e.atletas).length;
    if (n < 200) throw new Error('leitura veio com só ' + n + ' atletas; mantive o backup anterior');
    const sh = fedAba('FBERJ_Atletas', ['categoria', 'sigla', 'matricula', 'nome', 'entrouEm', 'vistoNaFBERJ']);
    const v = sh.getDataRange().getValues(); const idx = {};
    for (let i = 1; i < v.length; i++) idx[v[i][0] + '|' + v[i][1] + '|' + String(v[i][2])] = i + 1;
    const novos = []; const vistos = [];
    Object.keys(e.atletas).forEach(k => { const a = e.atletas[k]; if (idx[k]) vistos.push(idx[k]); else novos.push([a.cat, a.sig, "'" + a.reg, a.nome, agora, agora]); });
    if (novos.length) sh.getRange(sh.getLastRow() + 1, 1, novos.length, 6).setValues(novos);
    if (vistos.length && v.length > 1) { const col = sh.getRange(2, 6, v.length - 1, 1).getValues(); vistos.forEach(l => col[l - 2][0] = agora); sh.getRange(2, 6, v.length - 1, 1).setValues(col); }
    const eq = fedAba('FBERJ_Equipes', ['sigla', 'nome']); const ev = eq.getDataRange().getValues(); const tem = {};
    for (let i = 1; i < ev.length; i++) tem[ev[i][0]] = 1;
    const novEq = Object.keys(e.equipes).filter(s => !tem[s]).map(s => [s, e.equipes[s]]);
    if (novEq.length) eq.getRange(eq.getLastRow() + 1, 1, novEq.length, 2).setValues(novEq);
    det.push('Elenco: ' + n + ' atletas lidos, ' + novos.length + ' novos' + (novEq.length ? ', ' + novEq.length + ' equipes novas' : ''));
    // cestinhas: troca só se veio completo
    const sr = fedAba('FBERJ_Cestinhas', ['categoria', 'sigla', 'matricula', 'nome', 'total', 'tres', 'jogos', 'media']);
    const antR = Math.max(0, sr.getLastRow() - 1);
    if (e.pontos.resumo.length >= Math.max(200, antR * 0.8)) { if (antR) sr.getRange(2, 1, antR, 8).clearContent(); sr.getRange(2, 1, e.pontos.resumo.length, 8).setValues(e.pontos.resumo.map(x => [x[0], x[1], "'" + x[2], x[3], x[4], x[5], x[6], x[7]])); det.push('Cestinhas: ' + e.pontos.resumo.length); }
    else det.push('Cestinhas: leitura veio com ' + e.pontos.resumo.length + ' (antes ' + antR + '); mantive o anterior');
    const sj = fedAba('FBERJ_PontosJogo', ['categoria', 'fase', 'ocorrencia', 'sigla', 'adversario', 'matricula', 'pts']);
    const antJ = Math.max(0, sj.getLastRow() - 1);
    if (e.pontos.jogos.length >= Math.max(1000, antJ * 0.8)) { if (antJ) sj.getRange(2, 1, antJ, 7).clearContent(); sj.getRange(2, 1, e.pontos.jogos.length, 7).setValues(e.pontos.jogos.map(x => [x[0], x[1], x[2], x[3], x[4], "'" + x[5], x[6]])); det.push('Pontos por jogo: ' + e.pontos.jogos.length); }
    else det.push('Pontos por jogo: leitura veio com ' + e.pontos.jogos.length + ' (antes ' + antJ + '); mantive o anterior');
  } catch (err) { ok = false; det.push('Elenco: ERRO — ' + err.message); }
  // 2) tabela: troca só se veio completa
  try {
    const t = fedLerTabela(); const sh = fedAba('FBERJ_Tabela', ['categoria', 'fase', 'jogo', 'data', 'hora', 'timeA', 'timeB', 'ginasio', 'placarA', 'placarB', 'obs']);
    const antes = Math.max(0, sh.getLastRow() - 1);
    if (t.linhas.length < Math.max(50, antes * 0.8)) throw new Error('leitura veio com ' + t.linhas.length + ' jogos (antes ' + antes + '); mantive o backup anterior');
    if (antes) sh.getRange(2, 1, antes, 11).clearContent();
    sh.getRange(2, 1, t.linhas.length, 11).setNumberFormat('@').setValues(t.linhas);
    det.push('Tabela: ' + t.linhas.length + ' jogos' + (t.semSigla.length ? ' — nomes de time sem sigla: ' + t.semSigla.join(', ') : ''));
  } catch (err) { ok = false; det.push('Tabela: ERRO — ' + err.message); }
  // 3) tabela do feminino (3ª Taça Magaly Costa): troca só se veio completa
  try {
    const t = fedLerTabela(FED_TABELA_FEM_ID, true); const sh = fedAba('FBERJ_TabelaFem', ['categoria', 'fase', 'jogo', 'data', 'hora', 'timeA', 'timeB', 'ginasio', 'placarA', 'placarB', 'obs']);
    const antes = Math.max(0, sh.getLastRow() - 1);
    if (t.linhas.length < Math.max(5, antes * 0.8)) throw new Error('leitura veio com ' + t.linhas.length + ' jogos (antes ' + antes + '); mantive o backup anterior');
    if (antes) sh.getRange(2, 1, antes, 11).clearContent();
    sh.getRange(2, 1, t.linhas.length, 11).setNumberFormat('@').setValues(t.linhas);
    det.push('Feminino: ' + t.linhas.length + ' jogos');
  } catch (err) { ok = false; det.push('Feminino: ERRO — ' + err.message); }
  log.appendRow([agora, ok ? 'OK' : 'COM ERRO', det.join(' | ')]);
  fedLimparCache(); ['ces_n', 'sum_n', 'fem'].forEach(k => CacheService.getScriptCache().remove(k));
  return { ok: ok, detalhes: det };
}

function fedLimparCache() { const c = CacheService.getScriptCache(); const n = Number(c.get('fed_n') || 0); const ks = ['fed_n']; for (let i = 0; i < n; i++) ks.push('fed_' + i); c.removeAll(ks); }

function dadosFed() {
  const c = CacheService.getScriptCache(); const n = Number(c.get('fed_n') || 0);
  if (n) { const p = c.getAll(Array.from({ length: n }, (_, i) => 'fed_' + i)); const s = Array.from({ length: n }, (_, i) => p['fed_' + i]).join(''); if (s.length && Object.keys(p).length === n) return JSON.parse(s); }
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const sa = ss.getSheetByName('FBERJ_Atletas'), st = ss.getSheetByName('FBERJ_Tabela'), se = ss.getSheetByName('FBERJ_Equipes'), sl = ss.getSheetByName('FBERJ_Backup');
  if (!sa || !st) return { ok: false, erro: 'Backup ainda não feito.' };
  const elencos = {};
  sa.getDataRange().getDisplayValues().slice(1).forEach(r => { if (!r[0]) return; ((elencos[r[0]] = elencos[r[0]] || {})[r[1]] = elencos[r[0]][r[1]] || []).push([String(r[2]), r[3]]); });
  Object.keys(elencos).forEach(cat => Object.keys(elencos[cat]).forEach(s => elencos[cat][s].sort((a, b) => a[1].localeCompare(b[1], 'pt-BR'))));
  const equipes = {}; if (se) se.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) equipes[r[0]] = r[1]; });
  const tabela = [], placar = [];
  st.getDataRange().getDisplayValues().slice(1).forEach(r => { if (!r[0] || !r[5] || !r[6]) return; tabela.push([r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], (r[8] || r[9]) ? 1 : 0, r[10]]); placar.push([r[8], r[9]]); });
  let quando = ''; if (sl && sl.getLastRow() > 1) { const u = sl.getRange(sl.getLastRow(), 1, 1, 2).getValues()[0]; quando = Utilities.formatDate(new Date(u[0]), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') + ' (' + u[1] + ')'; }
  const res = { ok: true, atualizado: quando, elencos: elencos, equipes: equipes, tabela: tabela, placar: placar };
  try { const s = JSON.stringify(res); const parts = []; for (let i = 0; i < s.length; i += 90000) parts.push(s.slice(i, i + 90000)); const o = { fed_n: String(parts.length) }; parts.forEach((x, i) => o['fed_' + i] = x); c.putAll(o, 21600); } catch (e) {}
  return res;
}

function statusBackup(token) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  const sl = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('FBERJ_Backup'); if (!sl || sl.getLastRow() < 2) return { ok: true, ultimo: null };
  const v = sl.getRange(Math.max(2, sl.getLastRow() - 4), 1, Math.min(5, sl.getLastRow() - 1), 3).getValues().reverse();
  return { ok: true, ultimos: v.map(r => ({ quando: Utilities.formatDate(new Date(r[0]), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'), resultado: r[1], detalhes: r[2] })) };
}
function rodarBackupSite(token) { const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' }; return backupFBERJ(); }

/* Rodar UMA vez no editor: cria o gatilho diário das 20h e já faz o primeiro backup. */
function instalarBackupDiario() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backupFBERJ').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupFBERJ').timeBased().everyDays(1).atHour(20).inTimezone('America/Sao_Paulo').create();
  return backupFBERJ();
}


/* ---------- FBERJ — Mesa e Arbitragem (lista única) ----------
 * Aba FBERJ_Arbitragem: nome | funcao | nivel | origem (FBERJ/Klaus TV) | incluidoEm | incluidoPor
 * Nomes da FBERJ nunca são apagados; só os incluídos pelo site podem ser excluídos.
 */
const ARB_INICIAL = [["Mauricio Serour", "Árbitro(a)", "Especial"], ["Gabriel de Melo Alves", "Árbitro(a)", "Nacional"], ["Lucas Costa Maia da Silva", "Árbitro(a)", "Nacional"], ["Sandra Victória da Costa", "Árbitro(a)", "Nacional"], ["Tarik Silva Costa", "Árbitro(a)", "Nacional"], ["Juliane de Freitas Campos", "Árbitro(a)", "2ª"], ["Marco Aurelio Custodio Vieira", "Árbitro(a)", "2ª"], ["Rebeca Marciano Pockel Monteiro", "Árbitro(a)", "2ª"], ["Silvia Ferreira de Oliveira", "Árbitro(a)", "2ª"], ["Washington Luís Ferreira Franklin", "Árbitro(a)", "2ª"], ["Allan Madureira Ayres", "Árbitro(a)", "3ª"], ["Artur Louzada Guimarães", "Árbitro(a)", "3ª"], ["Flávia dos Santos Matias", "Árbitro(a)", "3ª"], ["Flávio Peixoto Luna", "Árbitro(a)", "3ª"], ["Felipe Miranda Rodrigues", "Árbitro(a)", "3ª"], ["Gabriela Aypana Delfino Braga", "Árbitro(a)", "3ª"], ["Gabriel Carvalho de Freitas", "Árbitro(a)", "3ª"], ["Gustavo Xisto Furtado", "Árbitro(a)", "3ª"], ["João Paulo Bernaldo de Paula", "Árbitro(a)", "3ª"], ["José Carlos Bastos de Azevedo", "Árbitro(a)", "3ª"], ["Luís Felipe Lage da Rocha", "Árbitro(a)", "3ª"], ["Luiz Felipe Cruz Soares", "Árbitro(a)", "3ª"], ["Maykel Bruno Gomes Fernandes", "Árbitro(a)", "3ª"], ["Michael dos Santos Braga", "Árbitro(a)", "3ª"], ["Rafael Stofel de Paulo", "Árbitro(a)", "3ª"], ["Saul Emmanuel de Mattos Quintans", "Árbitro(a)", "3ª"], ["Alvaro Francisco Gomes de Souza", "Árbitro(a)", "Estagiário(a)"], ["Hugo Ricardo Martins Ramalho", "Árbitro(a)", "Estagiário(a)"], ["Raphael de Miranda Machado", "Árbitro(a)", "Estagiário(a)"], ["Aline Miranda Vianna", "Oficial de mesa", "Especial e FIBA"], ["Elisângela Santana Cardoso", "Oficial de mesa", "Especial e FIBA"], ["Isabela Gomes Valentim", "Oficial de mesa", "Especial e FIBA"], ["Maria Claudia de Mattos Molina", "Oficial de mesa", "Especial e FIBA"], ["Marcelo Santana Cardoso", "Oficial de mesa", "Especial e FIBA"], ["Maria Thereza Rezende", "Oficial de mesa", "Especial e FIBA"], ["Rosângela Esteves Menezes", "Oficial de mesa", "Especial e FIBA"], ["Adriana Schimkus de Lima", "Oficial de mesa", "2ª"], ["Alessandra Rosa Lopes Zachi", "Oficial de mesa", "2ª"], ["Bárbara Silva de Oliveira", "Oficial de mesa", "2ª"], ["Denize de Lima Lourenço", "Oficial de mesa", "2ª"], ["Stéphanie Santana dos Santos Coelho", "Oficial de mesa", "2ª"], ["Alessandra de Oliveira Lages", "Oficial de mesa", "3ª"], ["Carolina Pontes Magalhães Cordeiro", "Oficial de mesa", "3ª"], ["Carolina Ribeiro Leite Chaves", "Oficial de mesa", "3ª"], ["Fernanda Rafaela Andrade Nunes", "Oficial de mesa", "3ª"], ["Helena La Marca de Azevedo", "Oficial de mesa", "3ª"], ["Igor Thales da Silva Souza", "Oficial de mesa", "3ª"], ["Jônatas Rodrigues Silva", "Oficial de mesa", "3ª"], ["Kátia Cristina de Lima Francisco", "Oficial de mesa", "3ª"], ["Luciano da Conceição Daniel", "Oficial de mesa", "3ª"], ["Mateus Pires Ribeiro", "Oficial de mesa", "3ª"], ["Milena Piva Rocha de Araujo", "Oficial de mesa", "3ª"], ["Raphael Queiroz Pereira Martins", "Oficial de mesa", "3ª"], ["Shirlei Regina de Oliveira", "Oficial de mesa", "3ª"], ["Yasmin Almeida Menzes de Souza", "Oficial de mesa", "3ª"], ["Elisangela Cristina Cardoso Mamede", "Oficial de mesa", "Estagiário(a)"], ["Lívia Cristina Alves Viana", "Oficial de mesa", "Estagiário(a)"], ["Mônica Meireles da Silva Molina Freitas", "Oficial de mesa", "Estagiário(a)"], ["Pablo Pereira Miranda da Silva", "Oficial de mesa", "Estagiário(a)"], ["Patrícia Filardi Monteiro", "Oficial de mesa", "Estagiário(a)"], ["Rafael de Lima Carvalho", "Oficial de mesa", "Estagiário(a)"], ["Vanessa de Oliveira Freitas", "Oficial de mesa", "Estagiário(a)"]];
function abaArbitragem() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID); let sh = ss.getSheetByName('FBERJ_Arbitragem');
  if (!sh) {
    sh = ss.insertSheet('FBERJ_Arbitragem'); sh.getRange(1, 1, 1, 6).setValues([['nome', 'funcao', 'nivel', 'origem', 'incluidoEm', 'incluidoPor']]); sh.setFrozenRows(1);
    const agora = new Date(); sh.getRange(2, 1, ARB_INICIAL.length, 6).setValues(ARB_INICIAL.map(r => [r[0], r[1], r[2], 'FBERJ', agora, 'PDF FBERJ']));
  }
  return sh;
}
function arbNorm(x) { return String(x || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLowerCase(); }
function lerArbitragem() {
  return abaArbitragem().getDataRange().getValues().slice(1).filter(r => String(r[0]).trim())
    .map((r, i) => ({ linha: i + 2, nome: String(r[0]).trim(), funcao: String(r[1] || ''), origem: String(r[3] || '') }));
}
function listarArbitragem(token) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  const l = lerArbitragem().sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
  return { ok: true, lista: l.map(x => ({ nome: x.nome, funcao: x.funcao, origem: x.origem })) };
}
function incluirArbitragem(token, nomes) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  const tem = {}; lerArbitragem().forEach(x => tem[arbNorm(x.nome)] = 1);
  const novos = []; (Array.isArray(nomes) ? nomes : [nomes]).forEach(x => {
    let n = (x && typeof x === 'object') ? x.nome : x; let f = (x && typeof x === 'object') ? arbFuncao(x.funcao) : '';
    n = String(n || '').replace(/\s+/g, ' ').trim(); if (n.length < 3 || n.length > 80) return;
    const k = arbNorm(n); if (tem[k]) return; tem[k] = 1; novos.push([n, f, '', 'Klaus TV', new Date(), u.usuario]);
  });
  if (novos.length) { const sh = abaArbitragem(); sh.getRange(sh.getLastRow() + 1, 1, novos.length, 6).setValues(novos); }
  return { ok: true, incluidos: novos.map(r => r[0]) };
}
function excluirArbitragem(token, nome) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  const x = lerArbitragem().find(r => arbNorm(r.nome) === arbNorm(nome));
  if (!x) return { ok: false, erro: 'Nome não encontrado.' };
  if (x.origem === 'FBERJ') return { ok: false, erro: 'Nome da lista oficial da FBERJ não pode ser excluído.' };
  abaArbitragem().deleteRow(x.linha); return { ok: true };
}

function arbFuncao(f) { f = arbNorm(f); return f.indexOf('arbitro') === 0 ? 'Árbitro(a)' : (f.indexOf('oficial') === 0 ? 'Oficial de mesa' : ''); }
function funcaoArbitragem(token, nome, funcao) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  const x = lerArbitragem().find(r => arbNorm(r.nome) === arbNorm(nome)); if (!x) return { ok: false, erro: 'Nome não encontrado.' };
  if (x.origem === 'FBERJ') return { ok: false, erro: 'Função da lista oficial da FBERJ não se altera.' };
  abaArbitragem().getRange(x.linha, 2).setValue(arbFuncao(funcao)); return { ok: true };
}


/* ================= SÚMULAS DA FBERJ (backup no nosso Drive) =================
 * Copia os PDFs da pasta CEMB 2026 da FBERJ para a pasta "Klaus TV - Súmulas FBERJ 2026" (link público só para ver).
 * Só copia o que é novo; nunca apaga. Se passar do tempo, agenda a continuação sozinho.
 * Aba FBERJ_Sumulas: origemId | categoria | fase | arquivo | copiaId | bytes | criadoNaFBERJ | copiadoEm
 */
const SUM_RAIZ = 'Klaus TV - Súmulas FBERJ 2026';
function sumPasta(nome, pai) {
  const it = pai ? pai.getFoldersByName(nome) : DriveApp.getFoldersByName(nome);
  if (it.hasNext()) return it.next();
  const f = pai ? pai.createFolder(nome) : DriveApp.createFolder(nome);
  if (!pai) { try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {} }
  return f;
}
function sumListar(paiId) {
  const out = []; let tok = null;
  do {
    const r = Drive.Files.list({ q: "'" + paiId + "' in parents and trashed = false", fields: 'nextPageToken, files(id, name, mimeType, size, createdTime)', pageSize: 1000, pageToken: tok, supportsAllDrives: true, includeItemsFromAllDrives: true });
    (r.files || []).forEach(f => out.push(f)); tok = r.nextPageToken;
  } while (tok);
  return out;
}
function backupSumulas() {
  const ini = Date.now(); const sh = fedAba('FBERJ_Sumulas', ['origemId', 'categoria', 'fase', 'arquivo', 'copiaId', 'bytes', 'criadoNaFBERJ', 'copiadoEm']);
  const feitos = {}; sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0] && r[4]) feitos[r[0]] = 1; });
  const raiz = sumPasta(SUM_RAIZ); const pend = [];
  // percorre categoria > fase > arquivos (e arquivos soltos em qualquer nível)
  const andar = (id, caminho) => sumListar(id).forEach(f => {
    if (f.mimeType === 'application/vnd.google-apps.folder') { if (caminho.length < 3) andar(f.id, caminho.concat([f.name])); }
    else if (!feitos[f.id]) pend.push({ f: f, caminho: caminho });
  });
  andar(FED_SUMULAS_ID, []);
  let n = 0; const novas = []; const pastas = {};
  for (const p of pend) {
    if (Date.now() - ini > 270000) break;
    const chave = p.caminho.join('/');
    let dest = pastas[chave]; if (!dest) { dest = raiz; p.caminho.forEach(nm => dest = sumPasta(nm, dest)); pastas[chave] = dest; }
    try {
      const c = Drive.Files.copy({ name: p.f.name, parents: [dest.getId()] }, p.f.id, { supportsAllDrives: true });
      novas.push([p.f.id, p.caminho[0] || '', p.caminho.slice(1).join(' / '), p.f.name, c.id, Number(p.f.size || 0), p.f.createdTime || '', new Date()]); n++;
    } catch (e) { novas.push([p.f.id, p.caminho[0] || '', p.caminho.slice(1).join(' / '), p.f.name, '', 0, p.f.createdTime || '', 'ERRO: ' + e.message]); }
    if (novas.length >= 25) { sh.getRange(sh.getLastRow() + 1, 1, novas.length, 8).setValues(novas); novas.length = 0; }
  }
  if (novas.length) sh.getRange(sh.getLastRow() + 1, 1, novas.length, 8).setValues(novas);
  const faltam = pend.length - n;
  // continua sozinho em 2 minutos se ainda faltar
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backupSumulasContinuar').forEach(t => ScriptApp.deleteTrigger(t));
  if (faltam > 0) ScriptApp.newTrigger('backupSumulasContinuar').timeBased().after(120000).create();
  if (n || faltam) fedAba('FBERJ_Backup', ['quando', 'resultado', 'detalhes']).appendRow([new Date(), faltam ? 'EM ANDAMENTO' : 'OK', 'Súmulas: ' + n + ' copiadas agora' + (faltam ? ', faltam ' + faltam : '')]);
  CacheService.getScriptCache().remove('sum_n');
  return { ok: true, copiadas: n, faltam: faltam, pasta: raiz.getUrl() };
}
function backupSumulasContinuar() { return backupSumulas(); }

/* Rodar UMA vez no editor: gatilhos diários (20h tabela/cestinhas/feminino, 21h súmulas) e primeira cópia. */
function instalarBackupSumulas() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backupSumulas').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupSumulas').timeBased().everyDays(1).atHour(21).inTimezone('America/Sao_Paulo').create();
  const r1 = backupFBERJ();
  const r2 = backupSumulas();
  Logger.log(JSON.stringify(r1)); Logger.log(JSON.stringify(r2));
  return r2;
}

/* ---------- dados públicos para as páginas Cestinhas e Súmulas (com cache) ---------- */
function cachePut(pref, obj, seg) { try { const s = JSON.stringify(obj); const parts = []; for (let i = 0; i < s.length; i += 90000) parts.push(s.slice(i, i + 90000)); const o = {}; o[pref + '_n'] = String(parts.length); parts.forEach((x, i) => o[pref + '_' + i] = x); CacheService.getScriptCache().putAll(o, seg); } catch (e) {} }
function cacheGet(pref) { const c = CacheService.getScriptCache(); const n = Number(c.get(pref + '_n') || 0); if (!n) return null; const p = c.getAll(Array.from({ length: n }, (_, i) => pref + '_' + i)); if (Object.keys(p).length !== n) return null; try { return JSON.parse(Array.from({ length: n }, (_, i) => p[pref + '_' + i]).join('')); } catch (e) { return null; } }
function fedQuando() { const sl = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('FBERJ_Backup'); if (!sl || sl.getLastRow() < 2) return ''; const v = sl.getRange(2, 1, sl.getLastRow() - 1, 2).getValues().filter(r => r[1] === 'OK' || r[1] === 'COM ERRO'); const u = v[v.length - 1]; return u ? Utilities.formatDate(new Date(u[0]), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') : ''; }

function dadosCestinhas() {
  const hit = cacheGet('ces'); if (hit) return hit;
  const ss = SpreadsheetApp.openById(PLANILHA_ID); const sr = ss.getSheetByName('FBERJ_Cestinhas');
  if (!sr || sr.getLastRow() < 2) return { ok: false, erro: 'Os cestinhas ainda não foram copiados da FBERJ.' };
  const se = ss.getSheetByName('FBERJ_Equipes'); const equipes = {}; if (se) se.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) equipes[r[0]] = r[1]; });
  const lista = sr.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => [r[0], r[1], String(r[2]), r[3], Number(r[4] || 0), Number(r[5] || 0), Number(r[6] || 0)]);
  const res = { ok: true, atualizado: fedQuando(), equipes: equipes, apelidos: mapaApelidos(), fotos: mapaFotos(), lista: lista };
  cachePut('ces', res, 3600); return res;
}

function dadosSumulas() {
  const hit = cacheGet('sum'); if (hit) return hit;
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const ssu = ss.getSheetByName('FBERJ_Sumulas'), st = ss.getSheetByName('FBERJ_Tabela'), sj = ss.getSheetByName('FBERJ_PontosJogo'), sa = ss.getSheetByName('FBERJ_Atletas');
  if (!ssu || ssu.getLastRow() < 2) return { ok: false, erro: 'As súmulas ainda estão sendo copiadas da FBERJ.' };
  // tabela por categoria + número do jogo
  const tab = {}; if (st) st.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0] && r[2]) tab[r[0] + '|' + r[2]] = r; });
  // nomes dos atletas
  const nomes = {}; if (sa) sa.getDataRange().getDisplayValues().slice(1).forEach(r => { nomes[r[0] + '|' + r[2]] = r[3]; });
  // pontos por jogo: cat|fase|sigla|adv -> {ocorrencia: [[reg, pts]]}
  const pj = {}; if (sj) sj.getDataRange().getValues().slice(1).forEach(r => { if (!r[0]) return; const k = r[0] + '|' + r[1] + '|' + r[3] + '|' + r[4]; ((pj[k] = pj[k] || {})[r[2]] = pj[k][r[2]] || []).push([String(r[5]), Number(r[6] || 0)]); });
  const cest = (cat, fase, sig, adv) => { const m = pj[cat + '|' + fase + '|' + sig + '|' + adv]; if (!m || Object.keys(m).length !== 1) return null; const l = m[Object.keys(m)[0]].slice().sort((a, b) => b[1] - a[1]); if (!l.length || !l[0][1]) return null; const top = l.filter(x => x[1] === l[0][1]).map(x => nomes[cat + '|' + sig + '|' + x[0]] || x[0]); return [top.join(' / '), l[0][1]]; };
  // estatísticas do Klaus (MVP e números do time), ligadas pelo nº do jogo
  const kt = {}; lerJogos().forEach(j => { const m = jsonOu(j.meta, {}); if (m.avulso) return; const num = String(j.jogoFed || '').split(' ')[0]; if (!num || j.status !== 'FINAL') return;
    const st2 = jsonOu(j.stats, {}); const atl = {}; ['A', 'B'].forEach(l => ((m['time' + l] || {}).atletas || []).forEach(a => atl[a.reg] = [a.nome, (m['time' + l] || {}).sigla]));
    let mvp = null; const time = {}; ['A', 'B'].forEach(l => { const sg = (m['time' + l] || {}).sigla; const t = { reb: 0, ast: 0, stl: 0, blk: 0, eff: 0 }; Object.keys(st2[l] || {}).forEach(reg => { const s = st2[l][reg]; ['reb', 'ast', 'stl', 'blk', 'eff'].forEach(k => t[k] += Number(s[k] || 0)); if (!mvp || Number(s.eff || 0) > mvp[1]) mvp = [(atl[reg] || [reg])[0], Number(s.eff || 0), sg]; }); time[sg] = t; });
    kt[j.categoria + '|' + num] = { id: j.id, mvp: mvp, time: time }; });
  const lista = ssu.getDataRange().getValues().slice(1).filter(r => r[4]).map(r => {
    const catP = String(r[1]).toUpperCase().replace(/\s+/g, ''); const cat = catP === 'ADULTO' ? 'Adulto' : (catP.match(/^SUB(\d+)/) ? 'Sub-' + catP.match(/^SUB(\d+)/)[1] : r[1]);
    const nm = String(r[3]); const num = (nm.match(/^\s*(?:jg|jogo)\.?\s*(\d+)/i) || [])[1] || '';
    const t = tab[cat + '|' + num]; const o = { cat: cat, pasta: String(r[2]), arquivo: nm.replace(/\.pdf$/i, '').trim(), id: r[4], num: num };
    if (t) { o.fase = t[1]; o.data = t[3]; o.hora = t[4]; o.a = t[5]; o.b = t[6]; o.pa = t[8]; o.pb = t[9]; o.gin = t[7];
      o.ca = cest(cat, t[1], t[5], t[6]); o.cb = cest(cat, t[1], t[6], t[5]); }
    const k = kt[cat + '|' + num]; if (k) o.kt = k;
    return o; });
  const se = ss.getSheetByName('FBERJ_Equipes'); const equipes = {}; if (se) se.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) equipes[r[0]] = r[1]; });
  // feminino (só tabela, ainda sem súmulas)
  const sf = ss.getSheetByName('FBERJ_TabelaFem'); const fem = sf ? sf.getDataRange().getDisplayValues().slice(1).filter(r => r[0] && r[5] && r[6]) : [];
  const res = { ok: true, atualizado: fedQuando(), equipes: equipes, lista: lista, feminino: fem };
  cachePut('sum', res, 1800); return res;
}
