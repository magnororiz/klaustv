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
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    switch (p.acao) {
      case 'login': return responder(login(p.usuario, p.senha));
      case 'trocarSenha': return responder(trocarSenha(p.token, p.senhaAtual, p.senhaNova));
      case 'eu': return responder(eu(p.token));
      case 'listarUsuarios': return responder(listarUsuarios(p.token));
      case 'criarUsuario': return responder(criarUsuario(p.token, p));
      case 'excluirUsuario': return responder(excluirUsuario(p.token, p.usuario));
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
function autorizar() { protegerSenhas(); Logger.log(MailApp.getRemainingDailyQuota() + ' e-mails disponíveis hoje'); }
