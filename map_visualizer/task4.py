import pandas as pd
from shapely.geometry import LineString, Point

routes_file = "routes.xlsx"
stops_file = "bus-stops.xlsx"

# Порог расстояния от остановки до маршрута (в метрах)
DISTANCE_THRESHOLD_METERS = 50
DEG_TO_M = 111139  # прибл. перевод градуса в метры


xls = pd.ExcelFile(routes_file)
route_lines = []

for sheet in xls.sheet_names:
    df = xls.parse(sheet)
    if {"latitude", "longitude"}.issubset(df.columns):
        df = df.dropna(subset=["latitude", "longitude"])
        df['route_id'] = sheet
        line = LineString(df[["longitude", "latitude"]].values)
        route_lines.append((sheet, line))

stops_df = pd.read_excel(stops_file)
stops_df = stops_df.dropna(subset=["latitude", "longitude"])

# === Привязка остановок к маршрутам по геометрии ===
matched_stops = []

for route_id, line in route_lines:
    for _, stop in stops_df.iterrows():
        stop_point = Point(stop["longitude"], stop["latitude"])
        distance_deg = line.distance(stop_point)
        distance_m = distance_deg * DEG_TO_M
        if distance_m <= DISTANCE_THRESHOLD_METERS:
            matched_stops.append({
                "route_id": route_id,
                "bus_stop_id": stop["bus_stop_id"],
                "bus_stop_name": stop["bus_stop_name"]
            })

matched_df = pd.DataFrame(matched_stops)

route_stops = matched_df.groupby("route_id")["bus_stop_id"].apply(set).to_dict()

from itertools import combinations

overlap_data = []

for (r1, s1), (r2, s2) in combinations(route_stops.items(), 2):
    common = s1 & s2
    n_common = len(common)
    if len(s1) == 0 or len(s2) == 0:
        continue
    p1 = round(100 * n_common / len(s1), 1)
    p2 = round(100 * n_common / len(s2), 1)
    duplicate = "да" if p1 >= 50 or p2 >= 50 else "нет"

    overlap_data.append({
        "Маршрут 1": r1,
        "Маршрут 2": r2,
        "Совпадающих остановок": n_common,
        "Совпадение % по маршруту 1": p1,
        "Совпадение % по маршруту 2": p2,
        "Дублируются": duplicate
    })


overlap_df = pd.DataFrame(overlap_data)
overlap_df.to_excel("дублирующиеся_маршруты.xlsx", index=False)
print("✅ Готово! Сохранено в файл: дублирующиеся_маршруты.xlsx")
