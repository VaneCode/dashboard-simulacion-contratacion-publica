import json, math, random, re, statistics, unicodedata
from collections import Counter, defaultdict
from datetime import date, timedelta
from pathlib import Path
import xlrd

ROOT=Path(__file__).resolve().parent
OUT=ROOT/'generated_data'; OUT.mkdir(exist_ok=True)
SRC=ROOT/'upload'/'PAC 2026 Guayaquil.xls'
SEED=2026

def txt(v):
    if v is None: return ''
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return re.sub(r'\s+',' ',str(v)).strip()

def key(s):
    return ''.join(c for c in unicodedata.normalize('NFD',txt(s).upper()) if unicodedata.category(c)!='Mn')

def money(v): return round(float(v or 0),2)
def iso(d): return d.isoformat() if d else ''

OBJECTIVES={
 'OBJ-01':('Fortalecer la gestión administrativa e institucional','Gestión administrativa'),
 'OBJ-02':('Mejorar la infraestructura y el equipamiento urbano','Infraestructura y obras'),
 'OBJ-03':('Optimizar los servicios urbanos y la gestión ambiental','Servicios urbanos y ambiente'),
 'OBJ-04':('Impulsar la transformación digital y la comunicación pública','Tecnología y comunicación'),
 'OBJ-05':('Fortalecer la inclusión y el desarrollo comunitario','Desarrollo social y comunitario'),
 'OBJ-06':('Mejorar la movilidad y la seguridad vial','Movilidad y seguridad vial'),
 'OBJ-07':('Fortalecer la salud, higiene y control sanitario','Salud e higiene'),
 'OBJ-08':('Fortalecer la gestión financiera, jurídica y de planificación','Gestión financiera y jurídica')}

def objective(desc, typ):
    d=key(desc)
    if key(typ)=='OBRAS': return 'OBJ-02'
    tests=[('OBJ-07',['MATADERO','CAMAL','HIGIENE','SANITAR','ALIMENTO']),('OBJ-03',['ASEO','DESECHO','AMBIENT','RESIDU','LIMPIEZA','PARQUE']),('OBJ-04',['SOFTWARE','SISTEMA','TECNOLOG','COMPUT','COMUNICACION','DATOS','LICENCIA']),('OBJ-06',['VEHIC','TRANSITO','MOVILIDAD','SENAL','VIAL','TRANSPORTE']),('OBJ-05',['SOCIAL','CULTUR','DEPORTE','COMUNIT','TURIS','EDUC']),('OBJ-08',['JURID','SEGURO','AUDITOR','FINAN','PLANIFIC','CONSULT'])]
    for code,words in tests:
        if any(w in d for w in words): return code
    return 'OBJ-01'

def load_pac():
    s=xlrd.open_workbook(str(SRC)).sheet_by_index(0)
    rows=[]
    for i in range(5,s.nrows):
        v=s.row_values(i)
        if not any(txt(x) for x in v): continue
        typ=key(v[3]).replace('CONSULTORIA','CONSULTORIA')
        q=1 if key(v[8])=='S' else 2 if key(v[9])=='S' else 3 if key(v[10])=='S' else 0
        qty=float(v[5] or 0); unit=float(v[7] or 0)
        rows.append({'id_pac':f'PAC-2026-{len(rows)+1:04d}','anio':int(v[0] or 2026),'ruc_entidad':'0960000220001','partida':txt(v[1]),'cpc':txt(v[2]),'tipo_compra':typ,'detalle':txt(v[4]),'cantidad':qty,'unidad':txt(v[6]),'costo_unitario':money(unit),'monto_planificado':money(qty*unit),'cuatrimestre':q,'tipo_producto':key(v[11]) or 'NO INFORMADO','catalogo_electronico':key(v[12]) or 'NO INFORMADO','procedimiento':key(v[13]) or 'NO INFORMADO','fondos_bid':key(v[14]) or 'NO','regimen':key(v[17]) or 'NO INFORMADO','tipo_presupuesto':key(v[18]) or 'NO INFORMADO'})
    return rows

def make_poa(pac):
    agg=defaultdict(lambda:{'monto_pac':0,'procesos':0})
    for r in pac:
        o=objective(r['detalle'],r['tipo_compra']); r['objetivo_sintetico']=o
        agg[o]['monto_pac']+=r['monto_planificado']; agg[o]['procesos']+=1
    poa=[]
    for idx,o in enumerate(OBJECTIVES,1):
        obj,act=OBJECTIVES[o]; a=agg[o]
        poa.append({'id_poa':f'POA-2026-{idx:02d}','objetivo_codigo':o,'objetivo_sintetico':obj,'actividad_sintetica':act,'indicador_sintetico':'Porcentaje de procesos vinculados con ejecución registrada','meta_sintetica_pct':85 if o!='OBJ-02' else 75,'presupuesto_sintetico':money(a['monto_pac']*1.03),'monto_pac_vinculado':money(a['monto_pac']),'procesos_pac':a['procesos'],'naturaleza':'SINTÉTICO — generado para fines académicos'})
    return poa

PARAMS={
 'CATALOGO ELECTRONICO':(.92,30),'INFIMA CUANTIA':(.88,25),'SUBASTA INVERSA ELECTRONICA':(.78,80),'LICITACION':(.70,135),'CONCURSO PUBLICO':(.68,120),'COMUNICACION SOCIAL CONTRATACION DIRECTA':(.80,65),'BIENES Y SERVICIOS UNICOS':(.82,60),'CONTRATOS ENTRE ENTIDADES PUBLICAS O SUBSIDIARIAS':(.85,55),'ARRENDAMIENTO DE BIENES INMUEBLES':(.90,45),'NO INFORMADO':(.72,90)}
SCENARIOS={'Conservador':(-.12,1.20),'Base':(0,1.0),'Favorable':(.10,.85)}

def simulate(pac,scenario='Base',seed=SEED,detail=True):
    rng=random.Random(seed); mod,durmod=SCENARIOS[scenario]; result=[]
    for r in pac:
        p,dur=PARAMS.get(r['procedimiento'],(.72,90)); p=max(.25,min(.98,p+mod))
        starts={1:(1,4),2:(5,8),3:(9,12),0:(1,12)}; lo,hi=starts[r['cuatrimestre']]
        month=rng.randint(lo,hi); planned=date(2026,month,rng.randint(1,24))
        prep=max(5,round(dur*durmod*rng.uniform(.12,.25))); proc=max(7,round(dur*durmod*rng.uniform(.30,.55))); execd=max(10,round(dur*durmod*rng.uniform(.40,.75)))
        publish=planned+timedelta(days=prep)
        success=rng.random()<=p
        cancelled=(not success and rng.random()<.20)
        award=publish+timedelta(days=proc) if success else None
        contract=award+timedelta(days=rng.randint(5,18)) if award else None
        finish=contract+timedelta(days=execd) if contract else None
        horizon=date(2026,12,31)
        if not success: status='CANCELADO' if cancelled else 'DESIERTO'
        elif finish<=horizon: status='FINALIZADO'
        elif contract<=horizon: status='EN EJECUCION'
        elif award<=horizon: status='ADJUDICADO'
        elif publish<=horizon: status='PUBLICADO'
        else: status='PLANIFICADO'
        comp=(-.02,.02) if r['procedimiento']=='CATALOGO ELECTRONICO' else (-.08,.04)
        awarded=money(r['monto_planificado']*(1+rng.uniform(*comp))) if success else 0
        executed=money(awarded*rng.uniform(.95,1.04)) if status=='FINALIZADO' else money(awarded*rng.uniform(.15,.85)) if status=='EN EJECUCION' else 0
        row={'id_proceso_sintetico':'PROC-'+r['id_pac'].split('-')[-1],'id_pac':r['id_pac'],'id_poa':f"POA-2026-{int(r['objetivo_sintetico'].split('-')[1]):02d}",'escenario':scenario,'procedimiento':r['procedimiento'],'monto_planificado':r['monto_planificado'],'fecha_planificada':iso(planned),'fecha_publicacion':iso(publish),'fecha_adjudicacion':iso(award),'fecha_contrato':iso(contract),'fecha_fin':iso(finish),'estado_simulado':status,'monto_adjudicado_sintetico':awarded,'monto_ejecutado_sintetico':executed,'dias_plan_adjudicacion':(award-planned).days if award else None,'naturaleza':'SINTÉTICO — semilla 2026'}
        result.append(row)
    return result

def metrics(cycle):
    n=len(cycle); completed=[x for x in cycle if x['estado_simulado']=='FINALIZADO']; success=[x for x in cycle if x['monto_adjudicado_sintetico']>0]; days=[x['dias_plan_adjudicacion'] for x in success if x['dias_plan_adjudicacion'] is not None]
    return {'registros':n,'tasa_materializacion_pct':round(100*len(success)/n,2),'tasa_finalizacion_pct':round(100*len(completed)/n,2),'desiertos_cancelados':sum(x['estado_simulado'] in ('DESIERTO','CANCELADO') for x in cycle),'monto_planificado':money(sum(x['monto_planificado'] for x in cycle)),'monto_adjudicado':money(sum(x['monto_adjudicado_sintetico'] for x in cycle)),'monto_ejecutado':money(sum(x['monto_ejecutado_sintetico'] for x in cycle)),'ahorro_adjudicacion_pct':round(100*(1-sum(x['monto_adjudicado_sintetico'] for x in success)/sum(x['monto_planificado'] for x in success)),2) if success else 0,'mediana_dias_adjudicacion':round(statistics.median(days),1) if days else 0}

def sensitivity(pac):
    rows=[]
    for sc in SCENARIOS:
        vals=[]
        for rep in range(1,101):
            m=metrics(simulate(pac,sc,SEED+rep,False)); m.update({'escenario':sc,'replicacion':rep}); rows.append(m); vals.append(m)
    summary=[]
    for sc in SCENARIOS:
        vs=[r for r in rows if r['escenario']==sc]
        rec={'escenario':sc}
        for field in ['tasa_materializacion_pct','tasa_finalizacion_pct','ahorro_adjudicacion_pct','mediana_dias_adjudicacion']:
            a=sorted(x[field] for x in vs); rec[field+'_p05']=a[4]; rec[field+'_mediana']=round(statistics.median(a),2); rec[field+'_p95']=a[94]
        summary.append(rec)
    return rows,summary

def main():
    pac=load_pac(); poa=make_poa(pac); cycle=simulate(pac); reps,sens=sensitivity(pac)
    counts={'tipo_compra':dict(Counter(x['tipo_compra'] for x in pac)),'procedimiento':dict(Counter(x['procedimiento'] for x in pac)),'cuatrimestre':{str(k):v for k,v in Counter(x['cuatrimestre'] for x in pac).items()},'estado_base':dict(Counter(x['estado_simulado'] for x in cycle))}
    dashboard={'fuente':'PAC 2026 Guayaquil','ruc':'0960000220001','registros_pac':len(pac),'monto_planificado':money(sum(x['monto_planificado'] for x in pac)),'objetivos_poa':len(poa),'metricas_base':metrics(cycle),'conteos':counts,'sensibilidad_resumen':sens}
    params=[{'procedimiento':k,'probabilidad_base':v[0],'duracion_referencia_dias':v[1]} for k,v in PARAMS.items()]
    dictionary=[{'campo':'monto_planificado','origen':'DERIVADO','definicion':'Cantidad anual multiplicada por costo unitario del PAC.'},{'campo':'objetivo_sintetico','origen':'SINTÉTICO','definicion':'Asignación académica por reglas de palabras clave y tipo de compra.'},{'campo':'fechas y estados simulados','origen':'SINTÉTICO','definicion':'Resultados reproducibles de simulación con semilla 2026.'},{'campo':'monto_adjudicado/ejecutado','origen':'SINTÉTICO','definicion':'Valores hipotéticos; no describen ejecución real del GAD.'}]
    for name,data in [('pac_limpio',pac),('poa_sintetico',poa),('ciclo_sintetico',cycle),('sensibilidad_replicaciones',reps),('sensibilidad_resumen',sens),('parametros',params),('diccionario',dictionary),('dashboard',dashboard)]:
        (OUT/f'{name}.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(dashboard,ensure_ascii=False,indent=2))

if __name__=='__main__': main()
