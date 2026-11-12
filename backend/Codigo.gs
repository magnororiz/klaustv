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
      case 'salvarStats': return responder(salvarStats(p.token, p.id, p.stats, p.status));
      case 'salvarFoto': return responder(salvarFoto(p.token, p.chave, p.imagem));
      case 'listarFotos': return responder(listarFotos(p.token));
      case 'salvarApelido': return responder(salvarApelido(p.token, p.reg, p.apelido));
      case 'aoVivoLista': return responder(aoVivoLista(p.categoria));
      case 'aoVivoJogo': return responder(aoVivoJogo(p.id));
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
  const meta = { timeA: j.timeA, timeB: j.timeB, oficiais: j.oficiais || {}, ginasio: j.ginasio || '', mandante: j.mandante || '', jogoFed: j.jogoFed || '', dataFed: j.dataFed || '' };
  const linha = [id, j.categoria, j.dataHora || '', j.timeA.sigla || '', j.timeB.sigla || '', j.ginasio || '', j.jogoFed || '', JSON.stringify(meta),
    atual ? atual.stats : '', atual ? atual.status : '', atual ? atual.placarA : 0, atual ? atual.placarB : 0, new Date(), u.usuario];
  if (atual) sh.getRange(atual.linha, 1, 1, COLS.length).setValues([linha]); else sh.appendRow(linha);
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
  abaJogos().deleteRow(j.linha); CacheService.getScriptCache().remove('lista');
  return { ok: true };
}
function salvarStats(token, id, stats, status) {
  const u = daSessao(token); if (!u) return { ok: false, erro: 'Sessão expirada. Entre de novo.' };
  const j = lerJogos().find(x => x.id === id); if (!j) return { ok: false, erro: 'Jogo não encontrado.' };
  const pl = placarDe(stats); const sh = abaJogos();
  sh.getRange(j.linha, 9, 1, 5).setValues([[JSON.stringify(stats || {}), status || j.status || '', pl[0], pl[1], new Date()]]);
  const c = CacheService.getScriptCache(); c.remove('j_' + id); c.remove('lista');
  return { ok: true, placar: pl };
}
function aoVivoLista(categoria) {
  const c = CacheService.getScriptCache(); let lista = jsonOu(c.get('lista'), null);
  if (!lista) {
    lista = lerJogos().map(x => ({ id: x.id, categoria: x.categoria, dataHora: x.dataHora instanceof Date ? x.dataHora.toISOString() : String(x.dataHora),
      timeA: x.timeA, timeB: x.timeB, placarA: Number(x.placarA || 0), placarB: Number(x.placarB || 0), status: x.status, ginasio: x.ginasio }));
    c.put('lista', JSON.stringify(lista), 20);
  }
  if (categoria) lista = lista.filter(x => x.categoria === categoria);
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
  return { ok: true, fotos: mapaFotos(), apelidos: mapaApelidos() };
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
  if (i > 0) { if (apelido) sh.getRange(i + 1, 2, 1, 2).setValues([[apelido, new Date()]]); else sh.deleteRow(i + 1); }
  else if (apelido) sh.appendRow([reg, apelido, new Date()]);
  CacheService.getScriptCache().remove('apelidos');
  return { ok: true };
}
