import { useState, useMemo } from "react";

const ROLES = {
  candidato:   { label:"Candidato",     icon:"👑", color:"#4A235A", bg:"#EDE7F6" },
  coordenador: { label:"Coordenador",   icon:"⭐", color:"#0C447C", bg:"#D6EAF8" },
  articulador: { label:"Articulador",   icon:"🔗", color:"#0F6E56", bg:"#C8F0DC" },
  juridico:    { label:"Jurídico",      icon:"⚖️", color:"#7B241C", bg:"#FDDCDC" },
  lider:       { label:"Líder de zona", icon:"📍", color:"#633806", bg:"#FFF3CD" },
  cabo:        { label:"Cabo eleitoral",icon:"👤", color:"#444441", bg:"#F1EFE8" },
  financeiro:  { label:"Financeiro",    icon:"💰", color:"#0C447C", bg:"#D6EAF8" },
  imprensa:    { label:"Imprensa",      icon:"📣", color:"#4A235A", bg:"#EDE7F6" },
};

const NOTIFS = {
  relatorio_diario: { label:"Relatório 20h",   icon:"📧" },
  alerta_critico:   { label:"Alerta crítico",  icon:"🚨" },
  ranking_semanal:  { label:"Ranking semanal", icon:"🏆" },
  cobranca_diaria:  { label:"Cobrança diária", icon:"💬" },
  spce_prazo:       { label:"Prazos SPCE",     icon:"📅" },
  resumo_semanal:   { label:"Resumo semanal",  icon:"📊" },
};

const ZONAS = ["Zona Norte","Zona Sul","Centro","Zona Oeste","Zona Leste"];

const INIT = [
  {id:"C1",nome:"Daniel Almeida",email:"daniel@campanha.com",whats:"(11)99000-0001",papel:"candidato",zona:"",bairro:"",notificacoes:["alerta_critico"],obs:"O candidato"},
  {id:"C2",nome:"Dulce Almeida",email:"dulce@campanha.com",whats:"(11)99000-0002",papel:"coordenador",zona:"",bairro:"",notificacoes:["relatorio_diario","alerta_critico"],obs:"Coordenadora geral"},
  {id:"C3",nome:"Carlos Menezes",email:"carlos@campanha.com",whats:"(11)99000-0003",papel:"lider",zona:"Zona Norte",bairro:"Santana",notificacoes:["cobranca_diaria","ranking_semanal"],obs:"Líder Zona Norte"},
  {id:"C4",nome:"Ana Rodrigues",email:"ana@campanha.com",whats:"(11)99000-0004",papel:"lider",zona:"Zona Sul",bairro:"Ipiranga",notificacoes:["cobranca_diaria","ranking_semanal"],obs:"Líder Zona Sul"},
  {id:"C5",nome:"João da Silva",email:"joao@campanha.com",whats:"(11)99000-0005",papel:"lider",zona:"Centro",bairro:"Sé",notificacoes:["cobranca_diaria"],obs:"Líder Centro"},
  {id:"C6",nome:"Roberto Almeida",email:"roberto@campanha.com",whats:"(11)99000-0006",papel:"lider",zona:"Zona Leste",bairro:"Penha",notificacoes:["cobranca_diaria"],obs:"Líder Zona Leste"},
  {id:"C7",nome:"Maria Conceição",email:"maria@campanha.com",whats:"(11)99000-0007",papel:"lider",zona:"Zona Oeste",bairro:"Lapa",notificacoes:["cobranca_diaria","ranking_semanal"],obs:"Líder Zona Oeste"},
  {id:"C8",nome:"Dr. Paulo Lima",email:"paulo@juridico.com",whats:"(11)99000-0008",papel:"juridico",zona:"",bairro:"",notificacoes:["spce_prazo"],obs:"Advogado eleitoral"},
];

const SCRIPT = `// Campanha Daniel Almeida 2026
// Cole em: Extensões → Apps Script

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Contatos");
  if (!ws) {
    ws = ss.insertSheet("Contatos");
    ws.getRange("A1:K1").setValues([[
      "ID","Timestamp","Nome","E-mail","WhatsApp",
      "Papel","Zona","Bairro","Notificações","Observações","Status"
    ]]).setBackground("#0D1F3C").setFontColor("#FFFFFF").setFontWeight("bold");
  }
  const d = JSON.parse(e.postData.contents);
  const id = "C" + ws.getLastRow();
  ws.appendRow([
    id, new Date(), d.nome, d.email, d.whats,
    d.papel, d.zona, d.bairro,
    d.notificacoes.join(", "), d.obs, "Ativo"
  ]);
  if (d.email) {
    MailApp.sendEmail({
      to: d.email,
      subject: "Bem-vindo(a) à Campanha Daniel Almeida 2026",
      htmlBody: bemVindoHTML(d.nome, d.papel, d.notificacoes)
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function dispararRelatorioDiario() {
  const emails = getEmailsPorNotif("relatorio_diario");
  emails.forEach(e => MailApp.sendEmail({
    to: e.email,
    subject: "Relatório diário — " + new Date().toLocaleDateString("pt-BR"),
    htmlBody: "<p>Olá " + e.nome + ", aqui está o relatório de hoje...</p>"
  }));
}

function getEmailsPorNotif(tipo) {
  const ws = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Contatos");
  if (!ws) return [];
  return ws.getRange(2, 1, ws.getLastRow()-1, 11).getValues()
    .filter(r => r[10]==="Ativo" && r[8].includes(tipo))
    .map(r => ({ nome: r[2], email: r[3] }));
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 Campanha Daniel")
    .addItem("Enviar relatório agora", "dispararRelatorioDiario")
    .addToUi();
}`;

const s = {
  wrap:  { fontFamily:"Arial, sans-serif", fontSize:14, color:"#1C1C1C", maxWidth:680 },
  topbar:{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:16, paddingBottom:12, borderBottom:"0.5px solid #E2E8F0" },
  tabBar:{ display:"flex", gap:2, marginBottom:16, borderBottom:"0.5px solid #E2E8F0" },
  card:  { background:"#F7F8FA", borderRadius:10, padding:16, marginBottom:12 },
  cardW: { background:"#FFFFFF", border:"0.5px solid #E2E8F0", borderRadius:10, padding:"10px 12px", marginBottom:8, display:"flex", alignItems:"center", gap:10 },
  inp:   { border:"0.5px solid #CBD5E0", borderRadius:6, padding:"7px 9px", fontSize:13, width:"100%", outline:"none", background:"#fff" },
  sel:   { border:"0.5px solid #CBD5E0", borderRadius:6, padding:"7px 9px", fontSize:13, outline:"none", background:"#fff" },
  btnP:  { background:"#1B4F8A", color:"#fff", border:"none", borderRadius:6, padding:"8px 16px", fontSize:13, fontWeight:500, cursor:"pointer" },
  btnS:  { background:"transparent", color:"#4A5568", border:"0.5px solid #CBD5E0", borderRadius:6, padding:"8px 12px", fontSize:13, cursor:"pointer" },
  sLabel:{ fontSize:11, fontWeight:500, color:"#718096", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:8, display:"block" },
  field: { display:"flex", flexDirection:"column", gap:4 },
  fLabel:{ fontSize:11, color:"#718096" },
  badge: { background:"#D6EAF8", color:"#1B4F8A", fontSize:11, fontWeight:500, padding:"2px 9px", borderRadius:12 },
};

function Avatar({ nome, papel }) {
  const ini = nome.split(" ").slice(0,2).map(p=>p[0]).join("").toUpperCase();
  const r = ROLES[papel] || ROLES.cabo;
  return <div style={{width:36,height:36,borderRadius:"50%",background:r.bg,color:r.color,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:500,flexShrink:0}}>{ini}</div>;
}

export default function App() {
  const [tab, setTab] = useState("cadastro");
  const [contatos, setContatos] = useState(INIT);
  const [form, setForm] = useState({ nome:"",email:"",whats:"",papel:"",zona:"",bairro:"",obs:"" });
  const [notifsSel, setNotifsSel] = useState([]);
  const [feedback, setFeedback] = useState(null);
  const [search, setSearch] = useState("");
  const [filterRole, setFilterRole] = useState("");
  const [filterZona, setFilterZona] = useState("");
  const [sheetsUrl, setSheetsUrl] = useState("https://docs.google.com/spreadsheets/d/1ABC123.../edit");
  const [connected, setConnected] = useState(true);
  const [copied, setCopied] = useState(false);

  const upd = k => e => setForm(f=>({...f,[k]:e.target.value}));

  const togNotif = n => setNotifsSel(s => s.includes(n)?s.filter(x=>x!==n):[...s,n]);

  const salvar = () => {
    if (!form.nome.trim() || !form.email.trim()) { setFeedback({type:"err",msg:"Nome e e-mail são obrigatórios."}); return; }
    if (!form.papel) { setFeedback({type:"err",msg:"Selecione o papel na campanha."}); return; }
    const novo = { ...form, id:"C"+(contatos.length+1), notificacoes:[...notifsSel] };
    setContatos(c=>[...c, novo]);
    setFeedback({type:"ok",msg:`✅ ${form.nome} cadastrado(a)! Adicionado ao Google Sheets · E-mail de boas-vindas enviado.`});
    setForm({nome:"",email:"",whats:"",papel:"",zona:"",bairro:"",obs:""});
    setNotifsSel([]);
    setTimeout(()=>setFeedback(null),4000);
  };

  const remover = id => setContatos(c=>c.filter(x=>x.id!==id));

  const lista = useMemo(()=> contatos.filter(c=>
    (!search||c.nome.toLowerCase().includes(search.toLowerCase())||c.email.toLowerCase().includes(search.toLowerCase()))&&
    (!filterRole||c.papel===filterRole)&&
    (!filterZona||c.zona===filterZona)
  ), [contatos,search,filterRole,filterZona]);

  const Tab = ({id,icon,label}) => (
    <button onClick={()=>setTab(id)} style={{...s.btnS,borderBottom:tab===id?"2px solid #1B4F8A":"2px solid transparent",borderLeft:"none",borderRight:"none",borderTop:"none",borderRadius:0,color:tab===id?"#1B4F8A":"#718096",fontWeight:tab===id?500:400,marginBottom:-1,padding:"8px 12px",fontSize:12}}>
      {icon} {label}
    </button>
  );

  const totalNotifs = useMemo(()=>{
    const m={};
    contatos.forEach(c=>c.notificacoes.forEach(n=>{m[n]=(m[n]||0)+1;}));
    return m;
  },[contatos]);

  return (
    <div style={s.wrap}>
      {/* topbar */}
      <div style={s.topbar}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:20}}>👥</span>
          <div>
            <div style={{fontSize:15,fontWeight:500}}>Cadastro de contatos</div>
            <div style={{fontSize:11,color:"#718096",marginTop:1}}>Campanha Daniel Almeida 2026</div>
          </div>
        </div>
        <span style={s.badge}>{contatos.length} contatos</span>
      </div>

      {/* tabs */}
      <div style={s.tabBar}>
        <Tab id="cadastro" icon="➕" label="Cadastrar"/>
        <Tab id="lista"    icon="📋" label="Contatos"/>
        <Tab id="sheets"   icon="🔗" label="Sheets"/>
        <Tab id="painel"   icon="📊" label="Painel"/>
      </div>

      {/* ── CADASTRAR ── */}
      {tab==="cadastro" && (
        <div>
          {feedback && <div style={{padding:"9px 12px",borderRadius:8,fontSize:12,marginBottom:12,background:feedback.type==="ok"?"#C8F0DC":"#FDDCDC",color:feedback.type==="ok"?"#0F6E56":"#7B241C",border:`0.5px solid ${feedback.type==="ok"?"#5DCAA5":"#F09595"}`}}>{feedback.msg}</div>}

          <div style={s.card}>
            <span style={s.sLabel}>Dados pessoais</span>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
              <div style={s.field}><label style={s.fLabel}>Nome completo *</label><input style={s.inp} placeholder="Ex: Carlos Menezes" value={form.nome} onChange={upd("nome")}/></div>
              <div style={s.field}><label style={s.fLabel}>E-mail *</label><input style={s.inp} type="email" placeholder="email@exemplo.com" value={form.email} onChange={upd("email")}/></div>
              <div style={s.field}><label style={s.fLabel}>WhatsApp</label><input style={s.inp} placeholder="(11) 99999-0000" value={form.whats} onChange={upd("whats")}/></div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              <div style={s.field}><label style={s.fLabel}>Zona eleitoral</label>
                <select style={s.sel} value={form.zona} onChange={upd("zona")}>
                  <option value="">— Selecionar —</option>
                  {ZONAS.map(z=><option key={z}>{z}</option>)}
                </select>
              </div>
              <div style={s.field}><label style={s.fLabel}>Bairro / Município</label><input style={s.inp} placeholder="Ex: Santana" value={form.bairro} onChange={upd("bairro")}/></div>
            </div>
          </div>

          <div style={s.card}>
            <span style={s.sLabel}>Papel na campanha *</span>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,marginBottom:2}}>
              {Object.entries(ROLES).map(([k,v])=>(
                <button key={k} onClick={()=>setForm(f=>({...f,papel:k}))} style={{padding:"8px 6px",borderRadius:8,border:form.papel===k?`1.5px solid ${v.color}`:"0.5px solid #E2E8F0",background:form.papel===k?v.bg:"#fff",cursor:"pointer",fontSize:11,fontWeight:form.papel===k?500:400,color:form.papel===k?v.color:"#718096",textAlign:"center"}}>
                  {v.icon} {v.label}
                </button>
              ))}
            </div>
          </div>

          <div style={s.card}>
            <span style={s.sLabel}>Notificações automáticas</span>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6,marginBottom:10}}>
              {Object.entries(NOTIFS).map(([k,v])=>(
                <button key={k} onClick={()=>togNotif(k)} style={{padding:"7px 6px",borderRadius:8,border:notifsSel.includes(k)?"1.5px solid #1A7A4A":"0.5px solid #E2E8F0",background:notifsSel.includes(k)?"#C8F0DC":"#fff",cursor:"pointer",fontSize:11,fontWeight:notifsSel.includes(k)?500:400,color:notifsSel.includes(k)?"#0F6E56":"#718096",display:"flex",alignItems:"center",justifyContent:"center",gap:4}}>
                  {v.icon} {v.label}
                </button>
              ))}
            </div>
            <div style={s.field}><label style={s.fLabel}>Observações</label><textarea style={{...s.inp,minHeight:56,resize:"vertical"}} placeholder="Ex: Responsável pelos cabos da Zona Norte" value={form.obs} onChange={upd("obs")}/></div>
          </div>

          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <button style={s.btnP} onClick={salvar}>💾 Salvar contato</button>
            <button style={s.btnS} onClick={()=>{setForm({nome:"",email:"",whats:"",papel:"",zona:"",bairro:"",obs:""});setNotifsSel([]);setFeedback(null);}}>Limpar</button>
          </div>
        </div>
      )}

      {/* ── LISTA ── */}
      {tab==="lista" && (
        <div>
          <div style={{display:"flex",gap:8,marginBottom:12}}>
            <input style={{...s.inp,flex:1}} placeholder="Buscar nome ou e-mail..." value={search} onChange={e=>setSearch(e.target.value)}/>
            <select style={{...s.sel,width:150}} value={filterRole} onChange={e=>setFilterRole(e.target.value)}>
              <option value="">Todos os papéis</option>
              {Object.entries(ROLES).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}
            </select>
            <select style={{...s.sel,width:130}} value={filterZona} onChange={e=>setFilterZona(e.target.value)}>
              <option value="">Todas as zonas</option>
              {ZONAS.map(z=><option key={z}>{z}</option>)}
            </select>
          </div>
          {lista.length===0 ? <div style={{textAlign:"center",padding:32,color:"#718096"}}>Nenhum contato encontrado.</div> : lista.map(c=>(
            <div key={c.id} style={s.cardW}>
              <Avatar nome={c.nome} papel={c.papel}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:13,fontWeight:500}}>{c.nome}</div>
                <div style={{fontSize:11,color:"#718096",marginTop:1}}>📧 {c.email}{c.whats?` · 📱 ${c.whats}`:""}</div>
                <div style={{display:"flex",gap:4,marginTop:5,flexWrap:"wrap"}}>
                  <span style={{fontSize:10,padding:"2px 7px",borderRadius:10,background:ROLES[c.papel]?.bg||"#F7F8FA",color:ROLES[c.papel]?.color||"#444"}}>{ROLES[c.papel]?.label||c.papel}</span>
                  {c.zona && <span style={{fontSize:10,padding:"2px 7px",borderRadius:10,background:"#F7F8FA",color:"#444"}}>{c.zona}</span>}
                  {c.notificacoes.slice(0,2).map(n=><span key={n} style={{fontSize:10,padding:"2px 7px",borderRadius:10,background:"#F7F8FA",color:"#718096"}}>{NOTIFS[n]?.icon} {NOTIFS[n]?.label||n}</span>)}
                </div>
              </div>
              <button onClick={()=>remover(c.id)} style={{background:"transparent",border:"0.5px solid #E2E8F0",borderRadius:6,width:28,height:28,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,color:"#718096"}}>🗑</button>
            </div>
          ))}
        </div>
      )}

      {/* ── SHEETS ── */}
      {tab==="sheets" && (
        <div>
          <div style={s.card}>
            <span style={s.sLabel}>URL do Google Sheets</span>
            <div style={{display:"flex",gap:8,marginBottom:10}}>
              <input style={{...s.inp,flex:1,fontSize:12}} value={sheetsUrl} onChange={e=>setSheetsUrl(e.target.value)} placeholder="https://docs.google.com/spreadsheets/d/..."/>
              <button style={s.btnP} onClick={()=>setConnected(true)}>🔌 Conectar</button>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12,marginBottom:10}}>
              <div style={{width:8,height:8,borderRadius:"50%",background:connected?"#1A7A4A":"#718096"}}/>
              <span style={{color:connected?"#1A7A4A":"#718096",fontWeight:500}}>{connected?"Planilha conectada":"Não conectada"}</span>
              {connected && <span style={{color:"#718096"}}>· Última sinc: agora</span>}
            </div>
            <div style={{fontSize:11,color:"#718096",lineHeight:1.6}}>Cada cadastro feito no app cria uma linha nova na aba <strong>Contatos</strong> do Sheets via Apps Script. E-mail de boas-vindas é enviado automaticamente.</div>
          </div>

          <div style={s.card}>
            <span style={s.sLabel}>Estrutura da aba "Contatos"</span>
            <div style={{background:"#fff",border:"0.5px solid #E2E8F0",borderRadius:8,padding:10,fontFamily:"monospace",fontSize:11,color:"#718096",lineHeight:1.9,marginBottom:10}}>
              <span style={{color:"#1C1C1C",fontWeight:500}}>A</span> ID &nbsp;·&nbsp; <span style={{color:"#1C1C1C",fontWeight:500}}>B</span> Timestamp &nbsp;·&nbsp; <span style={{color:"#1C1C1C",fontWeight:500}}>C</span> Nome &nbsp;·&nbsp; <span style={{color:"#1C1C1C",fontWeight:500}}>D</span> E-mail<br/>
              <span style={{color:"#1C1C1C",fontWeight:500}}>E</span> WhatsApp &nbsp;·&nbsp; <span style={{color:"#1C1C1C",fontWeight:500}}>F</span> Papel &nbsp;·&nbsp; <span style={{color:"#1C1C1C",fontWeight:500}}>G</span> Zona &nbsp;·&nbsp; <span style={{color:"#1C1C1C",fontWeight:500}}>H</span> Bairro<br/>
              <span style={{color:"#1C1C1C",fontWeight:500}}>I</span> Notificações &nbsp;·&nbsp; <span style={{color:"#1C1C1C",fontWeight:500}}>J</span> Observações &nbsp;·&nbsp; <span style={{color:"#1C1C1C",fontWeight:500}}>K</span> Status
            </div>
          </div>

          <div style={s.card}>
            <span style={s.sLabel}>Apps Script — cole no Google Sheets</span>
            <pre style={{background:"#fff",border:"0.5px solid #E2E8F0",borderRadius:8,padding:10,fontSize:10.5,fontFamily:"monospace",color:"#718096",maxHeight:180,overflowY:"auto",lineHeight:1.7,whiteSpace:"pre-wrap",marginBottom:10}}>{SCRIPT}</pre>
            <div style={{display:"flex",gap:8}}>
              <button style={s.btnP} onClick={()=>{navigator.clipboard?.writeText(SCRIPT).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000)});}}>
                {copied?"✅ Copiado!":"📋 Copiar código"}
              </button>
            </div>
          </div>

          <div style={s.card}>
            <span style={s.sLabel}>Passo a passo de instalação</span>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {[
                ["Abra sua planilha Google Sheets → Extensões → Apps Script"],
                ["Cole o código acima, salve com Ctrl+S — nomeie como Campanha Daniel"],
                ["Implantar → Nova implantação → App da Web → acesso: Qualquer pessoa"],
                ["Copie a URL de implantação e cole no campo acima para conectar"],
                ["Configure o gatilho: dispararRelatorioDiario → Por hora → 20h"],
              ].map(([t],i)=>(
                <div key={i} style={{display:"flex",gap:10,alignItems:"flex-start"}}>
                  <div style={{width:22,height:22,borderRadius:"50%",background:"#D6EAF8",color:"#1B4F8A",fontSize:11,fontWeight:500,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,marginTop:1}}>{i+1}</div>
                  <div style={{fontSize:12,lineHeight:1.5}}>{t}</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ── PAINEL ── */}
      {tab==="painel" && (
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
            {[
              {v:contatos.length,      l:"Total",          c:"#1B4F8A"},
              {v:contatos.filter(c=>c.papel==="lider").length, l:"Líderes de zona", c:"#856404"},
              {v:contatos.filter(c=>c.notificacoes.length>0).length, l:"Com notificação", c:"#1A7A4A"},
              {v:contatos.filter(c=>c.papel==="cabo").length, l:"Cabos eleitorais", c:"#718096"},
            ].map((st,i)=>(
              <div key={i} style={{background:"#F7F8FA",borderRadius:8,padding:12,textAlign:"center"}}>
                <div style={{fontSize:22,fontWeight:500,color:st.c}}>{st.v}</div>
                <div style={{fontSize:11,color:"#718096",marginTop:3}}>{st.l}</div>
              </div>
            ))}
          </div>

          <div style={{marginBottom:16}}>
            <span style={s.sLabel}>Contatos por zona</span>
            {ZONAS.map((z,i)=>{
              const n = contatos.filter(c=>c.zona===z).length;
              const max = Math.max(...ZONAS.map(zz=>contatos.filter(c=>c.zona===zz).length),1);
              const cors = ["#1B4F8A","#1A7A4A","#4A235A","#856404","#7B241C"];
              return (
                <div key={z} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                  <div style={{fontSize:12,width:100,flexShrink:0}}>{z}</div>
                  <div style={{flex:1,height:6,background:"#E2E8F0",borderRadius:3,overflow:"hidden"}}>
                    <div style={{width:`${Math.round(n/max*100)}%`,height:"100%",background:cors[i],borderRadius:3}}/>
                  </div>
                  <div style={{fontSize:11,color:"#718096",width:20,textAlign:"right"}}>{n}</div>
                </div>
              );
            })}
          </div>

          <div>
            <span style={s.sLabel}>Notificações mais ativas</span>
            {Object.entries(totalNotifs).sort((a,b)=>b[1]-a[1]).map(([n,cnt])=>{
              const max = Math.max(...Object.values(totalNotifs),1);
              return (
                <div key={n} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                  <div style={{fontSize:12,width:130,flexShrink:0}}>{NOTIFS[n]?.icon} {NOTIFS[n]?.label||n}</div>
                  <div style={{flex:1,height:5,background:"#E2E8F0",borderRadius:3,overflow:"hidden"}}>
                    <div style={{width:`${Math.round(cnt/max*100)}%`,height:"100%",background:"#1B4F8A",borderRadius:3}}/>
                  </div>
                  <div style={{fontSize:11,color:"#718096",width:16,textAlign:"right"}}>{cnt}</div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
