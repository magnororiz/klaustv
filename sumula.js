/* Klaus TV — PDF e imagem do jogo (usado pelo Scout e pelas Estatísticas ao vivo).
 * Monta a súmula a partir dos dados salvos do jogo (meta + stats), então serve para
 * qualquer jogo: em andamento, finalizado ou antigo.
 * Destaques: usa os escolhidos no Scout (stats.prem); os automáticos são recalculados;
 * jogo antigo sem destaques salvos ganha os 6 destaques calculados pelos números. */
(function(){
const ZERO={pts:0,twoMade:0,twoMiss:0,threeMade:0,threeMiss:0,ftMade:0,ftMiss:0,reb:0,ast:0,stl:0,blk:0,to:0,foul:0,eff:0};
const PREM_BASE=['MVP do jogo','Cestinha do jogo','Reboteiro do jogo','Defensor do jogo','Destaque do Time A','Destaque do Time B'];
const esc=x=>String(x==null?'':x).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const V=(p,c)=>+p.s[c]||0;
const DEF=p=>V(p,'reb')+V(p,'stl')+V(p,'blk');
// mesmos critérios do Scout
const CRIT={0:[p=>V(p,'eff'),p=>V(p,'pts'),p=>V(p,'reb'),DEF],1:[p=>V(p,'pts'),p=>V(p,'eff')],2:[p=>V(p,'reb'),DEF,p=>V(p,'eff'),p=>V(p,'pts')],3:[DEF,p=>V(p,'stl')+V(p,'blk'),p=>V(p,'eff'),p=>V(p,'pts')],
  4:[p=>V(p,'eff'),p=>V(p,'pts')],5:[p=>V(p,'eff'),p=>V(p,'pts')]};
function ordena(js,cr){ return js.filter(p=>cr[0](p)>0).sort((a,b)=>{ for(const f of cr){ const d=f(b)-f(a); if(d) return d; } return 0; }); }
function pctTxt(m,e){ const t=m+e; return m+'/'+t+' ('+(t?Math.round(m*100/t):0)+'%)'; }
function dataTxt(dh){ const d=String(dh||'').replace(' ','T'); if(!/^\d{4}-\d{2}-\d{2}/.test(d)) return ''; return d.slice(8,10)+'/'+d.slice(5,7)+'/'+d.slice(0,4)+(d.length>15?' às '+d.slice(11,16):''); }

function curto(n){ const w=String(n||'').trim().split(/\s+/).filter(Boolean); return w.length>1?w[0]+' '+w[w.length-1]:(w[0]||''); }
function jogadores(jogo,o){ const m=jogo.meta||{}, st=jogo.stats||{}, out=[];
  ['A','B'].forEach(t=>{ const tm=m['time'+t]||{}; const S=st[t]||{};
    (tm.atletas||[]).forEach(a=>{ const s=Object.assign({},ZERO,S[a.reg]||{}); const ap=(o.apelidos&&o.apelidos[a.reg])||a.apelido||'';
      out.push({reg:String(a.reg),t,tn:o.nomeTime(t),num:a.numero||'',nome:curto(a.nome)+(ap?' ('+ap+')':''),s}); }); });
  return out; }

function destaques(jogo,js,o){ const pr=(jogo.stats||{}).prem; const out=[];
  const nomeP=i=>i<6?PREM_BASE[i].replace('Time A',o.nomeTime('A')).replace('Time B',o.nomeTime('B')):'';
  const auto=i=>{ if(i===4||i===5){ const t=i===4?'A':'B'; return ordena(js.filter(p=>p.t===t),CRIT[i]).slice(0,1); }
    const r=ordena(js,CRIT[i]); if(!r.length) return []; if(i!==1) return [r[0]];
    const mx=V(r[0],'pts'); return r.filter(p=>V(p,'pts')===mx); }; // empate na cestinha: todos
  if(Array.isArray(pr)&&pr.length){
    pr.forEach(x=>{ const i=+x[0]; const nome=i<6?nomeP(i):(x[3]||''); if(!nome) return;
      if(i<4 && x[2]!=='1'){ auto(i).forEach(pl=>out.push({p:nome,pl})); return; }
      const pl=js.find(p=>p.reg===String(x[4]||'')); if(pl) out.push({p:nome,pl}); });
  } else [0,1,2,3,4,5].forEach(i=>auto(i).forEach(pl=>out.push({p:nomeP(i),pl})));
  return out; }

const CSS='#ktSumBox .stJ{display:inline-block;margin-top:6px;padding:3px 12px;border-radius:12px;background:#fff3cd;color:#8a6100;font-size:12px;font-weight:900;letter-spacing:1px}#ktSumBox .stJ.fin{background:#e7e9ee;color:#333}#ktSumBox table.qz{width:auto;margin:14px auto 4px;border-collapse:collapse;font-size:14px;table-layout:auto}#ktSumBox table.qz th,#ktSumBox table.qz td{border:1px solid #c9d3e3;padding:5px 14px;text-align:center;width:auto!important;min-width:44px}#ktSumBox table.qz th{background:#1e5bd8;color:#fff}#ktSumBox table.qz .l{text-align:left;font-weight:700}#ktSumBox{width:1000px;background:#fff;color:#111;font-family:Arial,sans-serif;padding:26px 30px;box-sizing:border-box}#ktSumBox .cab{display:flex;align-items:center;gap:18px;border-bottom:4px solid #1e5bd8;padding-bottom:14px}#ktSumBox .cab img{height:96px}#ktSumBox .pl{flex:1;text-align:center}#ktSumBox .pl .t{display:flex;align-items:center;justify-content:center;gap:18px;font-weight:900;font-size:24px}#ktSumBox .pl .n{font-size:54px;color:#1e5bd8;font-family:Oswald,Arial}#ktSumBox .pl .x{font-size:26px;color:#999}#ktSumBox .inf{color:#555;font-size:14px;margin-top:6px}#ktSumBox .sT{background:#1e5bd8;color:#fff;font-weight:800;font-size:17px;padding:7px 12px;margin-top:18px;border-radius:6px 6px 0 0}#ktSumBox table{width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed}#ktSumBox th.cN{width:32px}#ktSumBox th.cJ{width:220px}#ktSumBox th.cA{width:84px}#ktSumBox th{background:#e8eefb;color:#1a2c55;padding:6px 3px;font-size:12px}#ktSumBox td{padding:5px 3px;text-align:center;border-bottom:1px solid #e4e7ec;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}#ktSumBox .l{text-align:left}#ktSumBox tr.tot td{background:#0d1a33;color:#ffd54f;font-weight:800}#ktSumBox .prem{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;padding:10px;background:#fff8e1;border:1px solid #ffd54f;border-top:none}#ktSumBox .prem div{display:flex;flex-direction:column;background:#fff;border-radius:8px;padding:8px 10px;border:1px solid #f0e2a8}#ktSumBox .prem span{font-size:12px;color:#8a6d00;font-weight:700;text-transform:uppercase}#ktSumBox .prem b{font-size:15px}#ktSumBox .prem small{color:#666}#ktSumBox .rod{margin-top:18px;border-top:2px solid #e4e7ec;padding-top:10px;display:flex;justify-content:space-between;color:#555;font-size:13px}';

// jogo = {meta, stats, categoria, dataHora, ginasio, jogoFed, status}; o = {nomeTime(lado), apelidos, logo}
function montar(jogo,o){
  const m=jogo.meta||{}, tn=o.nomeTime;
  const js=jogadores(jogo,o); const tot=t=>js.filter(p=>p.t===t).reduce((a,p)=>a+V(p,'pts'),0);
  const data=dataTxt(jogo.dataHora)||new Date().toLocaleDateString('pt-BR');
  const tab=t=>{ const ps=js.filter(p=>p.t===t).sort((a,b)=>V(b,'pts')-V(a,'pts')||V(b,'eff')-V(a,'eff'));
    const sum=k=>ps.reduce((a,p)=>a+V(p,k),0);
    return '<div class="sT">'+esc(tn(t))+'</div><table><tr><th class="cN">#</th><th class="cJ l">Jogador</th><th>PTS</th><th class="cA">2P</th><th class="cA">3P</th><th class="cA">LL</th><th>REB</th><th>AST</th><th>ROU</th><th>TOC</th><th>ERR</th><th>FP</th><th>EF</th></tr>'+
    ps.map(p=>'<tr><td>'+esc(p.num)+'</td><td class="l">'+esc(p.nome)+'</td><td><b>'+V(p,'pts')+'</b></td><td>'+pctTxt(V(p,'twoMade'),V(p,'twoMiss'))+'</td><td>'+pctTxt(V(p,'threeMade'),V(p,'threeMiss'))+'</td><td>'+pctTxt(V(p,'ftMade'),V(p,'ftMiss'))+'</td><td>'+V(p,'reb')+'</td><td>'+V(p,'ast')+'</td><td>'+V(p,'stl')+'</td><td>'+V(p,'blk')+'</td><td>'+V(p,'to')+'</td><td>'+V(p,'foul')+'</td><td><b>'+V(p,'eff')+'</b></td></tr>').join('')+
    '<tr class="tot"><td colspan="2" class="l">Total</td><td>'+sum('pts')+'</td><td>'+pctTxt(sum('twoMade'),sum('twoMiss'))+'</td><td>'+pctTxt(sum('threeMade'),sum('threeMiss'))+'</td><td>'+pctTxt(sum('ftMade'),sum('ftMiss'))+'</td><td>'+sum('reb')+'</td><td>'+sum('ast')+'</td><td>'+sum('stl')+'</td><td>'+sum('blk')+'</td><td>'+sum('to')+'</td><td>'+sum('foul')+'</td><td>'+sum('eff')+'</td></tr></table>'; };
  const prem=destaques(jogo,js,o);
  const premHtml=prem.length?'<div class="sT">🏆 Destaques do jogo</div><div class="prem">'+prem.map(x=>'<div><span>'+esc(x.p)+'</span><b>'+esc(x.pl.nome)+'</b><small>'+esc(x.pl.tn)+'</small></div>').join('')+'</div>':'';
  const info=[jogo.categoria, data, jogo.ginasio||m.ginasio, (jogo.jogoFed||m.jogoFed)?'Jogo FBERJ nº '+(jogo.jogoFed||m.jogoFed):''].filter(Boolean).map(esc).join(' · ');
  const fin=jogo.status==='FINAL';
  const per=(jogo.stats||{}).periodos||{};
  const qv=(t,q)=>{ const v=(per[t]||[])[q-1]; return (v===undefined||v===null||v==='')?'':String(v); };
  const QQ=(qv('A',5)!==''||qv('B',5)!=='')?[1,2,3,4,5]:[1,2,3,4];
  const temQ=QQ.some(q=>qv('A',q)!==''||qv('B',q)!=='');
  const qzHtml=temQ?'<table class="qz"><tr><th class="l">'+(fin?'Placar final':'Placar parcial')+'</th>'+QQ.map(q=>'<th>'+(q===5?'OT':q+'º Q')+'</th>').join('')+'<th>Final</th></tr>'+['A','B'].map(t=>'<tr><td class="l">'+esc(tn(t))+'</td>'+QQ.map(q=>'<td>'+(qv(t,q)===''?'—':esc(qv(t,q)))+'</td>').join('')+'<td><b>'+tot(t)+'</b></td></tr>').join('')+'</table>':'';
  let box=document.getElementById('ktSumula'); if(!box){ box=document.createElement('div'); box.id='ktSumula'; document.body.appendChild(box); }
  box.style.cssText='position:fixed;left:-12000px;top:0;';
  box.innerHTML='<div id="ktSumBox"><style>'+CSS+'</style>'+
   '<div class="cab"><img src="'+(o.logo||'logo_klaustv_tv.png')+'" alt=""><div class="pl"><div class="t"><span>'+esc(tn('A'))+'</span><span class="n">'+tot('A')+'</span><span class="x">x</span><span class="n">'+tot('B')+'</span><span>'+esc(tn('B'))+'</span></div><div class="inf">'+info+'</div><div class="stJ'+(fin?' fin':'')+'">'+(fin?'JOGO FINALIZADO':'JOGO EM ANDAMENTO · placar parcial')+'</div></div></div>'+qzHtml+
   tab('A')+tab('B')+premHtml+'<div class="rod"><span>📸 @basquete.rj</span><span>💬 (21) 99809-2305</span><span>🌐 klaustv.com</span></div></div>';
  return {box, el:document.getElementById('ktSumBox'), nome:('KlausTV '+tn('A')+' x '+tn('B')+' '+data.slice(0,10)).replace(/[\/\\:*?"<>|]+/g,'-')};
}
function lib(src){ return new Promise((ok,erro)=>{ if(document.querySelector('script[src="'+src+'"]')) return ok(); const e=document.createElement('script'); e.src=src; e.onload=ok; e.onerror=erro; document.head.appendChild(e); }); }
async function exportar(tipo,jogo,o){
  try{ await lib('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
    if(tipo==='pdf') await lib('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  }catch(e){ alert('Não foi possível carregar o gerador. Confira a internet e tente de novo.'); return false; }
  const r=montar(jogo,o);
  try{ const imgs=[...r.el.querySelectorAll('img')]; await Promise.all(imgs.map(i=>i.complete?0:new Promise(ok=>{ i.onload=i.onerror=ok; })));
    const cv=await html2canvas(r.el,{scale:2,backgroundColor:'#ffffff',useCORS:true});
    if(tipo==='img'){
      await new Promise(fim=>cv.toBlob(async b=>{ const f=new File([b],r.nome+'.png',{type:'image/png'});
        if(navigator.canShare && navigator.canShare({files:[f]}) && /Android|iPhone|iPad/i.test(navigator.userAgent)){ try{ await navigator.share({files:[f],title:r.nome}); return fim(); }catch(e){} }
        const a=document.createElement('a'); a.href=URL.createObjectURL(b); a.download=r.nome+'.png'; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),4000); fim(); },'image/png'));
    } else {
      const {jsPDF}=window.jspdf; const pdf=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
      const W=200, H=cv.height*W/cv.width, pg=285; const img=cv.toDataURL('image/jpeg',0.92);
      if(H<=pg) pdf.addImage(img,'JPEG',5,5,W,H);
      else { let y=0; while(y<H){ pdf.addImage(img,'JPEG',5,5-y,W,H); y+=pg; if(y<H) pdf.addPage(); } }
      pdf.save(r.nome+'.pdf');
    }
  } finally { r.box.innerHTML=''; }
  return true;
}
window.KTSumula={montar,exportar,pctTxt};
})();
