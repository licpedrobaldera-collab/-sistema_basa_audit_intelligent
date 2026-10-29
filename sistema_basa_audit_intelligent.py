# -*- coding: utf-8 -*-
"""
BASA - AUDIT INTELLIGENT USA - V4 BASA FULL - ACTUALIZADO
Lic. Pedro A. Baldera Rondón, CPA - BASA
GitHub: basa-audit-intelligent - Streamlit Cloud Online
Master: basa.admin / BASA_2026_MASTER - ID interno oculto
Cliente: Solo ID / ITIN / NSS - 1 dato + Contrato Legal SHA-256
"""
import io, json, hashlib
from pathlib import Path
from datetime import datetime
import pandas as pd
import openpyxl
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
try:
    import fitz
    import streamlit as st
    import plotly.graph_objects as go
    import plotly.express as px
    PLOTLY=True
except:
    PLOTLY=False

MASTER_USER="basa.admin"
MASTER_PASS="BASA_2026_MASTER"
MASTER_ID_INT="05700084394"
MASTER_NOMBRE="BASA - Lic. Pedro A. Baldera Rondón, CPA"
BASE=Path(__file__).parent
EVID=BASE/"evidencias_auditoria"; EVID.mkdir(exist_ok=True)
REP=BASE/"reportes_exportados"; REP.mkdir(exist_ok=True)
USERS_FILE=BASE/"usuarios_basa.json"
HIST_FILE=BASE/"historico_basa.json"
CONT_FILE=BASE/"contratos_firmados.json"

def load_users():
    if USERS_FILE.exists():
        try: return json.loads(USERS_FILE.read_text(encoding='utf-8'))
        except: pass
    return {MASTER_USER:{"usuario":MASTER_USER,"clave":MASTER_PASS,"id_interno":MASTER_ID_INT,"nombre":MASTER_NOMBRE,"rol":"MASTER_ADMIN","activo":True,"historial_claves":[MASTER_PASS],"creado":datetime.now().isoformat()}}
def save_users(u): USERS_FILE.write_text(json.dumps(u,indent=2,ensure_ascii=False),encoding='utf-8')
def load_hist():
    if HIST_FILE.exists():
        try: return json.loads(HIST_FILE.read_text(encoding='utf-8'))
        except: return []
    return []
def save_hist(d): HIST_FILE.write_text(json.dumps(d,indent=2,ensure_ascii=False),encoding='utf-8')
def load_cont():
    if CONT_FILE.exists():
        try: return json.loads(CONT_FILE.read_text(encoding='utf-8'))
        except: return []
    return []
def save_cont(d): CONT_FILE.write_text(json.dumps(d,indent=2,ensure_ascii=False),encoding='utf-8')

class Extractor:
    @staticmethod
    def docx_extract(b):
        doc=docx.Document(b); txt=[p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for t in doc.tables:
            for r in t.rows:
                rt=" | ".join([c.text.strip() for c in r.cells if c.text.strip()])
                if rt: txt.append(rt)
        return "\n".join(txt)
    @staticmethod
    def excel_extract(b):
        xls=pd.ExcelFile(b); lines=[]
        for sh in xls.sheet_names:
            df=pd.read_excel(xls,sheet_name=sh); lines.append(f"=== {sh} ===\n{df.to_string(index=False)}")
        return "\n".join(lines)
    @staticmethod
    def pdf_extract(b):
        doc=fitz.open(stream=b,filetype="pdf"); return "\n".join([doc[p].get_text("text") for p in range(len(doc)) if doc[p].get_text("text").strip()])
    @classmethod
    def process(cls,up):
        name=up.name.lower(); content=up.read(); up.seek(0)
        try:
            if name.endswith((".docx",".doc")): return cls.docx_extract(io.BytesIO(content))
            elif name.endswith((".xlsx",".xls",".csv")): return cls.excel_extract(io.BytesIO(content))
            elif name.endswith(".pdf"): return cls.pdf_extract(content)
            else: return content.decode("utf-8",errors="ignore")
        except Exception as e: return f"Error {name}: {e}"

DEFAULT_RULES=[
    {"id":"FUS-01","rubro":"Gobernanza","titulo":"Deficiencias NOBACI","kw":["nobaci","ética"]},
    {"id":"FUS-02","rubro":"Tecnología","titulo":"PETI y SIGEF","kw":["peti","sigef"]},
    {"id":"FUS-03","rubro":"RRHH","titulo":"Talento Humano","kw":["novedades","expedientes"]},
    {"id":"FUS-04","rubro":"Nóminas","titulo":"Nóminas Internas","kw":["nómina"]},
    {"id":"FUS-05","rubro":"Ética","titulo":"Nepotismo","kw":["declaración"]},
]
def classify_f(t):
    tl=t.lower()
    for r in DEFAULT_RULES:
        for kw in r["kw"]:
            if kw in tl: return r["id"]
    return "FUS-01"

def build_excel(entity,findings):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Matriz BASA"
    ws.merge_cells("A1:H1"); ws["A1"]=f"{entity.upper()} - BASA AUDIT INTELLIGENT - {MASTER_NOMBRE}"
    hdr=["No.","Pág","Rubro","Título","Condición","Monto","Fusión","Score"]
    ws.append([]); ws.append(hdr)
    for idx,it in enumerate(findings,start=5):
        ws.cell(row=idx,column=1,value=it.get("id","")); ws.cell(row=idx,column=2,value=it.get("pagina","Pág.10"))
        ws.cell(row=idx,column=3,value=it.get("rubro","")); ws.cell(row=idx,column=4,value=it.get("titulo",""))
        ws.cell(row=idx,column=5,value=it.get("descripcion","")); cm=ws.cell(row=idx,column=6,value=it.get("monto",0)); cm.number_format='"RD$"#,##0.00'
        ws.cell(row=idx,column=7,value=classify_f(it.get("titulo",""))); ws.cell(row=idx,column=8,value=it.get("scoreCalidad",8))
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def build_word(entity,audit_type,period,auditor,findings):
    doc=docx.Document(); p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f"{entity.upper()}\nBASA AUDIT INTELLIGENT\n"); r.bold=True; r.font.size=Pt(14)
    pm=doc.add_paragraph(); pm.add_run("BASA: ").bold=True; pm.add_run(f"{MASTER_NOMBRE}\n")
    pm.add_run("Auditor: ").bold=True; pm.add_run(f"{auditor}\n")
    doc.add_paragraph(f"Contrato Legal Firmado - Info Real Verificable - SHA-256: {hashlib.sha256(str(findings).encode()).hexdigest()[:16]}")
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf

def render_dashboard(findings,currency="RD$"):
    if not PLOTLY or not findings: return
    df=pd.DataFrame(findings)
    if 'scoreCalidad' not in df.columns: df['scoreCalidad']=8.0
    if 'monto' not in df.columns: df['monto']=0
    grouped=df.groupby('rubro').agg(count=('id','count'),avg_riesgo=('scoreCalidad','mean'),max_riesgo=('scoreCalidad','max'),monto=('monto','sum')).reset_index()
    st.markdown("---"); st.subheader("📊 Dashboard BASA - Radar 0-10 + Presupuesto")
    c1,c2=st.columns(2)
    with c1:
        fig=go.Figure()
        fig.add_trace(go.Scatterpolar(r=grouped['avg_riesgo'].tolist()+[grouped['avg_riesgo'].iloc[0]],theta=grouped['rubro'].tolist()+[grouped['rubro'].iloc[0]],fill='toself',name='Promedio'))
        fig.add_trace(go.Scatterpolar(r=grouped['max_riesgo'].tolist()+[grouped['max_riesgo'].iloc[0]],theta=grouped['rubro'].tolist()+[grouped['rubro'].iloc[0]],fill='toself',name='Máx',opacity=0.4))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True,range=[0,10])),height=350); st.plotly_chart(fig,use_container_width=True)
    with c2:
        fig2=px.bar(grouped.sort_values('monto'),x='monto',y='rubro',orientation='h',color='count',text='monto'); fig2.update_layout(height=350); st.plotly_chart(fig2,use_container_width=True)

def run_app():
    st.set_page_config(page_title="BASA Audit Intelligent - V4",layout="wide")
    users=load_users()
    if 'user_logged' not in st.session_state: st.session_state.user_logged=None
    if 'contrato' not in st.session_state: st.session_state.contrato=False

    if not st.session_state.user_logged:
        st.title("🛡️ BASA - Audit Intelligent USA - V4 BASA ACTUALIZADO")
        st.markdown(f"**{MASTER_NOMBRE} - Plataforma GitHub Online**"); st.markdown("---")
        c1,c2=st.columns(2)
        with c1:
            st.subheader("🔐 Acceso BASA - Usuario y Clave")
            u=st.text_input("Usuario BASA",placeholder="basa.admin"); p=st.text_input("Contraseña",type="password",placeholder="Tu clave BASA")
            if st.button("Entrar BASA",type="primary"):
                if u in users and users[u]["clave"]==p and users[u]["activo"]: st.session_state.user_logged=users[u]; st.rerun()
                else: st.error("Usuario o clave incorrecta")
        with c2:
            st.subheader("👤 Cliente - Solo 1 Dato: ID / ITIN / NSS")
            idc=st.text_input("ID / ITIN / NSS",placeholder="Ej: 123-45-6789"); nomc=st.text_input("Nombre Completo"); telc=st.text_input("Tel / WhatsApp")
            if st.button("Entrar como Cliente - Cargar Soportes"):
                if idc and nomc:
                    cid=f"CLIENTE_{idc.replace('-','').replace(' ','')}"
                    if cid not in users:
                        users[cid]={"usuario":cid,"clave":idc,"id_interno":idc,"nombre":nomc,"rol":"CLIENTE","activo":True,"tel":telc,"historial_claves":[idc],"creado":datetime.now().isoformat()}; save_users(users)
                    st.session_state.user_logged=users[cid]; st.rerun()
                else: st.warning("Coloca ID/ITIN/NSS y Nombre")
        st.stop()

    user=st.session_state.user_logged
    st.sidebar.success(f"👤 {user['nombre']} | Usuario: {user['usuario']} | Rol: {user['rol']}")
    with st.sidebar.expander("🔑 Cambiar Mi Clave BASA"):
        act=st.text_input("Clave Actual",type="password",key="act"); nueva=st.text_input("Nueva Clave",type="password",key="new"); nueva2=st.text_input("Repetir Nueva Clave",type="password",key="new2")
        if st.button("Cambiar Clave y Guardar Histórico"):
            if act==user["clave"] and nueva==nueva2 and nueva:
                users[user["usuario"]]["clave"]=nueva; users[user["usuario"]]["historial_claves"].append(nueva); save_users(users); st.success("Clave cambiada"); st.session_state.user_logged=users[user["usuario"]]
            else: st.error("Clave actual incorrecta o no coinciden")
    if st.sidebar.button("Cerrar Sesión BASA"): st.session_state.user_logged=None; st.session_state.contrato=False; st.rerun()

    if not st.session_state.contrato and user['rol']!="MASTER_ADMIN":
        st.title("📜 Contrato Legal BASA - Responsabilidad Real"); st.warning("Debe firmar contrato legal para continuar")
        st.markdown(f"Yo, **{user['nombre']}** ID **{user['id_interno']}**, declaro bajo juramento que toda información y evidencias cargadas son **reales, verídicas y verificables** y soy **responsable legal**. Autorizo a BASA validar con IRS.")
        nf=st.text_input("Nombre para firma digital",value=user['nombre']); c1=st.checkbox("✅ Declaro info REAL y VERIFICABLE - Responsable legal"); c2=st.checkbox("✅ Autorizo validación IRS y acepto términos BASA"); c3=st.checkbox("✅ Entiendo penalidades por info falsa")
        if st.button("✍️ Firmar Contrato Legal y Continuar",type="primary"):
            if c1 and c2 and c3 and nf:
                cont={"usuario":user['usuario'],"nombre":nf,"id":user['id_interno'],"fecha":datetime.now().isoformat(),"hash":hashlib.sha256(f"{user['usuario']}{nf}{datetime.now().isoformat()}".encode()).hexdigest()}
                cs=load_cont(); cs.append(cont); save_cont(cs); st.session_state.contrato=True; st.success(f"Firmado Hash: {cont['hash'][:16]}"); st.rerun()
            else: st.error("Marque 3 casillas y nombre")
        st.stop()
    else:
        if user['rol']!="MASTER_ADMIN": st.sidebar.success("✅ Contrato Legal Firmado")

    st.sidebar.header("⚙️ Config BASA"); entity=st.sidebar.text_input("Entidad/Cliente","Hospital Materno Infantil"); audit_type=st.sidebar.selectbox("Modalidad",["Forense Antifraude","Financiera Especial","Integral","Obras","Salud"]); period=st.sidebar.text_input("Período","2020-2021"); auditor=st.sidebar.text_input("Auditor",f"{MASTER_NOMBRE}"); pais=st.sidebar.selectbox("País",["🇩🇴 RD - RD$","🇺🇸 USA - USD $","🇪🇸 ES - EUR €","🇲🇽 MX"])

    st.title("🛡️ BASA - Audit Intelligent - V4 BASA ACTUALIZADO"); st.markdown(f"**BASA - {MASTER_NOMBRE} - GitHub Online - Data 100% Real**")
    if 'findings' not in st.session_state:
        st.session_state.findings=[{"id":1,"pagina":"Pág.10","titulo":"Deficiencias registro novedades","rubro":"RRHH","monto":0,"descripcion":"Falta BD","scoreCalidad":7.2},{"id":32,"pagina":"Pág.28","titulo":"Pagos nómina interna","rubro":"Nóminas","monto":18509628,"descripcion":"Sin políticas","scoreCalidad":5.5}]

    uploaded=st.file_uploader("📂 Cargar Soportes y Evidencias Ley (W2, 1099, ID, PDF, Excel, Imágenes)",type=["docx","doc","xlsx","xls","csv","pdf","txt","png","jpg","jpeg","webp"],accept_multiple_files=True)
    if uploaded:
        texts=[]
        for f in uploaded:
            txt=Extractor.process(f); (EVID/f.name).write_bytes(f.getvalue()); texts.append(f"{f.name}: {txt[:800]}")
        st.text_area("Texto Extraído Real BASA", "\n\n".join(texts)[:4000],height=150); st.success(f"{len(uploaded)} evidencias guardadas en {EVID}")

    st.subheader(f"📋 Matriz Hallazgos BASA - Cliente: {entity}")
    if st.session_state.findings:
        df=pd.DataFrame(st.session_state.findings); edited=st.data_editor(df,num_rows="dynamic",use_container_width=True)
        c1,c2=st.columns(2)
        with c1:
            if st.button("💾 Guardar Histórico BASA",type="primary"): st.session_state.findings=edited.to_dict('records'); st.success("Guardado BASA")
        with c2:
            if st.button("🧹 Base Limpia (0) BASA"): st.session_state.findings=[]; st.rerun()
    else:
        st.info("Base Limpia (0) BASA"); 
        if st.button("🔄 Restaurar Catálogo BASA"): st.session_state.findings=[{"id":1,"pagina":"Pág.10","titulo":"Deficiencias","rubro":"RRHH","monto":0,"descripcion":"Falta","scoreCalidad":7.2}]; st.rerun()

    st.subheader("🔥 Matriz Riesgos BASA")
    if st.session_state.findings: st.dataframe(pd.DataFrame(st.session_state.findings)[['id','pagina','rubro','titulo','monto','scoreCalidad']].style.background_gradient(cmap='RdYlGn',subset=['scoreCalidad']),use_container_width=True)
    render_dashboard(st.session_state.findings,pais.split("-")[-1].strip())

    st.subheader("📚 Histórico BASA"); hist=load_hist(); cn=st.text_input("Nombre caso BASA",f"Caso {entity} {datetime.now().strftime('%Y-%m-%d')}")
    if st.button("💾 Guardar Caso Histórico BASA"):
        nuevo={"codigo":f"BASA-{datetime.now().strftime('%Y%m%d%H%M%S')}","nombre":cn,"entidad":entity,"auditor":user['usuario'],"fecha":datetime.now().isoformat(),"findings":st.session_state.findings,"monto":sum(f.get('monto',0) for f in st.session_state.findings)}; hist.append(nuevo); save_hist(hist); st.success(f"Guardado {nuevo['codigo']}")
    if hist:
        sel=st.selectbox("Retomar caso",[f"{h['codigo']} - {h['nombre']} - {h['entidad']}" for h in hist])
        if st.button("📂 Retomar BASA"):
            idx=[f"{h['codigo']} - {h['nombre']} - {h['entidad']}" for h in hist].index(sel); st.session_state.findings=hist[idx]['findings']; st.rerun()

    st.markdown("---"); st.subheader("📤 Exportar BASA + GitHub")
    if st.button("🚀 Generar Informes BASA - Excel/Word",type="primary"):
        if st.session_state.findings:
            eb=build_excel(entity,st.session_state.findings); wb=build_word(entity,audit_type,period,auditor,st.session_state.findings)
            (REP/f"Matriz_BASA_{entity.replace(' ','_')}.xlsx").write_bytes(eb.getvalue()); (REP/f"Informe_BASA_{entity.replace(' ','_')}.docx").write_bytes(wb.getvalue())
            c1,c2=st.columns(2)
            with c1: st.download_button("📊 Excel BASA",data=eb,file_name=f"Matriz_BASA_{entity}.xlsx")
            with c2: st.download_button("📄 Word BASA + Contrato",data=wb,file_name=f"Informe_BASA_{entity}.docx")
            st.success(f"Reportes en {REP} - Listo GitHub"); st.code(f"git add .; git commit -m 'BASA {entity}'; git push origin main",language="bash")
        else: st.warning("Base Limpia")

    st.markdown("---"); st.subheader("🌐 GitHub Audit Intelligent USA")
    st.markdown("1. Crea repo `basa-audit-intelligent` en GitHub\n2. Sube este archivo + requirements.txt\n3. Deploy en https://share.streamlit.io → New App → Te da link https://basa-audit-intelligent.streamlit.app")
    st.link_button("🌐 Crear Repo GitHub BASA", "https://github.com/new")

if __name__=="__main__":
    run_app()
