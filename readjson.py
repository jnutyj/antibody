# read_json_fields.py
import json
from pathlib import Path
import numpy as np

def read_json_fields(json_path: str):
    """
    读取指定 JSON 中的字段：
      - "imgt_frm_identity"
      - "Abnormal Cys"
      - "Free Cysteine"
      - "Glycosylation"
      - "Deamidation"
      - "Isomerization"
      - "id"

    返回一个字典（字段存在则取值，否则为 None）。
    """
    path = Path(json_path)
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    #print(data[0])
    #fields = [
    #    "imgt_frm_identity",
    #    "Abnormal Cys",
    #    "Free Cysteine",
    #    "Glycosylation",
    #    "Deamidation",
    #    "Isomerization",
    #    "id",
    #]

    #result = {k: None for k in fields}

    ## 常见情况：顶层就是一个 dict，字段直接在顶层
    #if isinstance(data, dict):
    #    for k in fields:
    #        if k in data:
    #            result[k] = data.get(k)

    # 若顶层不是 dict（例如 list），这里按“尝试取第一个元素为 dict”的宽松方式处理
    #elif isinstance(data, list) and data and isinstance(data[0], dict):
    #    first = data[0]
    #    for k in fields:
    #        if k in first:
    #            result[k] = first.get(k)

    return data


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python read_json_fields.py yourfile.json")
        sys.exit(1)

    out = read_json_fields(sys.argv[1])
    
    #kk=json.loads(out[0]['domain_color'])
    #for j in range(len(kk)):
    #   print(kk[j])
    #print(json.loads(out[0]['CDR3_json'])[-1][0][0])
    
    #exit(0)
    for i in range(len(out)):
       #kk=json.loads(out[i]['domain_color'])
       #for j in range(len(kk)):
       #   print(kk[j])    
       #    if r
       if 'red' in out[i]['domain_color']:
           #print(out[i]['domain_color'].split())
           cdr1_0=json.loads(out[i]['CDR1_json'])[0][0][0]
           cdr1_1=json.loads(out[i]['CDR1_json'])[-1][0][0]
           cdr2_0=json.loads(out[i]['CDR2_json'])[0][0][0]
           cdr2_1=json.loads(out[i]['CDR2_json'])[-1][0][0]
           cdr3_0=json.loads(out[i]['CDR3_json'])[0][0][0]
           cdr3_1=json.loads(out[i]['CDR3_json'])[-1][0][0] 
           print(out[i]['id'], out[i]["imgt_frm_identity"],out[i]["Abnormal Cys"],out[i]["Free Cysteine"],out[i]["Glycosylation"],out[i]["Deamidation"],out[i]["Isomerization"] , "CDR1:%i-%i"%(cdr1_0,cdr1_1), "CDR2:%i-%i"%(cdr2_0,cdr2_1), "CDR3:%i-%i"%(cdr3_0,cdr3_1))
       else:
           print(out[i]['id'], out[i]["imgt_frm_identity"])

