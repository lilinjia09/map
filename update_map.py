import pandas as pd
import requests
import json
import os
import time

EXCEL_FILE = "class_info.xlsx"
OUTPUT_FILE = "data.js"
CACHE_FILE = "city_cache.json"

def get_coord(city, cache):
    if city in cache:
        return cache[city]

    url = "https://nominatim.openstreetmap.org/search"
    params = {"q": city, "format": "json", "limit": 1}
    try:
        # Nominatim 需要 User-Agent，且限制频率
        r = requests.get(url, params=params, headers={"User-Agent":"StudentMapScript/1.0"}, timeout=10)
        data = r.json()
        if data:
            coord = [float(data[0]["lon"]), float(data[0]["lat"])]
            cache[city] = coord
            print(f"成功获取坐标: {city} -> {coord}")
            return coord
        else:
            print(f"查询不到城市: {city}")
            return None
    except Exception as e:
        print(f"请求出错: {city}, 错误: {e}")
        return None

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"错误：找不到文件 {EXCEL_FILE}")
        return

    # 加载缓存
    cache = {}
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            cache = json.load(f)

    # --- 关键修改点 1: header=1 表示使用 Excel 的第二行作为表头 ---
    try:
        df = pd.read_excel(EXCEL_FILE, header=1)
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return

    # 去除列名两端的空格，防止匹配失败
    df.columns = [str(c).strip() for c in df.columns]
    print("识别到的列名:", df.columns.tolist())

    result = []

    for index, row in df.iterrows():
        # --- 关键修改点 2: 修正对应的列名 ---
        name = str(row.get("姓名", "")).strip()
        city = str(row.get("城市", "")).strip()
        university = str(row.get("大学", "")).strip()  # Excel 里叫 大学
        major = str(row.get("专业", "")).strip()      # Excel 里叫 专业
        words = str(row.get("寄语", "")).strip()      # Excel 里叫 寄语
        photo = str(row.get("头像", "")).strip()      # Excel 里叫 头像

        # 排除空数据 (处理 NaN)
        if not name or name == "nan" or not city or city == "nan":
            continue

        # 获取坐标
        coord = get_coord(city, cache)
        if not coord:
            continue

        # 处理头像为空的情况
        if photo == "" or photo == "nan":
            photo = f"https://ui-avatars.com/api/?name={name}"

        result.append({
            "name": name,
            "value": coord,
            "ext": {
                "job": university,  # 对应原代码的学校
                "ind": major,       # 对应原代码的专业
                "words": words,
                "photo": photo
            }
        })
        
        # 频率限制
        if city not in cache:
            time.sleep(1.1)

    # 保存缓存
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    # 保存为 JS 文件
    js_content = "var rawData = " + json.dumps(result, ensure_ascii=False, indent=2) + ";"
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(js_content)

    print(f"\n处理完成！共生成 {len(result)} 条数据。")

if __name__ == "__main__":
    main()