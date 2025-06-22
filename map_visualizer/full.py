import pandas as pd
import numpy as np
from shapely.geometry import LineString, Point
from geopy.distance import geodesic
from itertools import combinations
import warnings

# === Параметры ===
routes_file = "routes.xlsx"
stops_file = "bus-stops.xlsx"
DISTANCE_THRESHOLD_METERS = 50
DEG_TO_M = 111139
ZIGZAG_ANGLE_THRESHOLD = 60
ZIGZAG_COUNT_THRESHOLD = 5
LOOP_DISTANCE_THRESHOLD = 500

# === Загрузка маршрутов ===
xls = pd.ExcelFile(routes_file)
all_routes = []

for sheet in xls.sheet_names:
    df = xls.parse(sheet)
    if {'latitude', 'longitude'}.issubset(df.columns):
        df = df.dropna(subset=['latitude', 'longitude'])
        df['route_id'] = sheet
        if 'type' not in df.columns:
            df['type'] = 'forward'
        all_routes.append(df)

routes_df = pd.concat(all_routes, ignore_index=True)

# === Загрузка остановок ===
stops_df = pd.read_excel(stops_file).dropna(subset=['latitude', 'longitude'])

# === Привязка остановок к маршрутам ===
matched = []

for (route_id, direction), group in routes_df.groupby(['route_id', 'type']):
    line = LineString(group[['longitude', 'latitude']].values)
    for _, stop in stops_df.iterrows():
        stop_point = Point(stop['longitude'], stop['latitude'])
        distance_deg = line.distance(stop_point)
        distance_m = distance_deg * DEG_TO_M
        if distance_m <= DISTANCE_THRESHOLD_METERS:
            matched.append({
                "route_id": route_id,
                "direction": direction,
                "bus_stop_id": stop["bus_stop_id"],
                "bus_stop_name": stop["bus_stop_name"],
                "latitude": stop["latitude"],
                "longitude": stop["longitude"]
            })

matched_df = pd.DataFrame(matched)

# === Пересадочные узлы ===
stop_routes = {}
for _, row in matched_df.iterrows():
    stop_id = row["bus_stop_id"]
    route = row["route_id"]
    stop_name = row["bus_stop_name"]
    if stop_id not in stop_routes:
        stop_routes[stop_id] = {"bus_stop_name": stop_name, "routes": set()}
    stop_routes[stop_id]["routes"].add(route)

hubs = []
for stop_id, info in stop_routes.items():
    hubs.append({
        "bus_stop_id": stop_id,
        "bus_stop_name": info["bus_stop_name"],
        "routes_count": len(info["routes"]),
        "routes": sorted(list(info["routes"]))
    })

df_hubs = pd.DataFrame(hubs).sort_values("routes_count", ascending=False)
df_hubs.to_excel("пересадочные_узлы.xlsx", index=False)

# === Дублирующиеся маршруты ===
route_stops = matched_df.groupby("route_id")["bus_stop_id"].apply(set).to_dict()
overlap_data = []

for (r1, s1), (r2, s2) in combinations(route_stops.items(), 2):
    common = s1 & s2
    if not s1 or not s2:
        continue
    p1 = round(100 * len(common) / len(s1), 1)
    p2 = round(100 * len(common) / len(s2), 1)
    overlap_data.append({
        "Маршрут 1": r1,
        "Маршрут 2": r2,
        "Совпадающих остановок": len(common),
        "Совпадение % по маршруту 1": p1,
        "Совпадение % по маршруту 2": p2,
        "Дублируются": "да" if p1 >= 50 or p2 >= 50 else "нет"
    })

df_overlap = pd.DataFrame(overlap_data)
df_overlap.to_excel("дублирующиеся_маршруты.xlsx", index=False)

# === Анализ формы маршрута: зигзаг, кольцо, линейный ===
def angle_between(p1, p2, p3):
    a = np.array(p1)
    b = np.array(p2)
    c = np.array(p3)
    ab = a - b
    cb = c - b
    cosine_angle = np.dot(ab, cb) / (np.linalg.norm(ab) * np.linalg.norm(cb))
    angle_rad = np.arccos(np.clip(cosine_angle, -1.0, 1.0))
    return np.degrees(angle_rad)

form_results = []
for (route_id, direction), group in routes_df.groupby(['route_id', 'type']):
    coords = list(zip(group['latitude'], group['longitude']))
    if len(coords) < 3:
        form_results.append({
            "route_id": route_id,
            "direction": direction,
            "тип_маршрута": "недостаточно данных",
            "комментарий": "меньше 3 точек"
        })
        continue

    start = coords[0]
    end = coords[-1]
    loop_dist = geodesic(start, end).meters

    sharp_turns = 0
    for i in range(1, len(coords) - 1):
        angle = angle_between(coords[i - 1], coords[i], coords[i + 1])
        if angle < ZIGZAG_ANGLE_THRESHOLD:
            sharp_turns += 1

    if loop_dist <= LOOP_DISTANCE_THRESHOLD:
        typ = "кольцевой"
        comment = f"расстояние между концами {int(loop_dist)} м"
    elif sharp_turns >= ZIGZAG_COUNT_THRESHOLD:
        typ = "зигзагообразный"
        comment = f"{sharp_turns} острых углов < {ZIGZAG_ANGLE_THRESHOLD}°"
    else:
        typ = "линейный"
        comment = f"{sharp_turns} поворотов, расстояние концов {int(loop_dist)} м"

    form_results.append({
        "route_id": route_id,
        "direction": direction,
        "тип_маршрута": typ,
        "комментарий": comment
    })

df_form = pd.DataFrame(form_results)
df_form.to_excel("тип_маршрутов.xlsx", index=False)

print("✅ Все файлы сохранены:")
print("- пересадочные_узлы.xlsx")
print("- дублирующиеся_маршруты.xlsx")
print("- тип_маршрутов.xlsx")
