# read_json_fields.py
import json
from pathlib import Path


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
    #print(out)
    for i in range(len(out)):
       if 'red' in out[i]['domain_color']:
           #print(out[i]['domain_color'].split())
           print(out[i]['id'], out[i]["imgt_frm_identity"],out[i]["Abnormal Cys"],out[i]["Free Cysteine"],out[i]["Glycosylation"],out[i]["Deamidation"],out[i]["Isomerization"])
       else:
           print(out[i]['id'], out[i]["imgt_frm_identity"])

