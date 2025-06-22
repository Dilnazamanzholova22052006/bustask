import pandas as pd
from shapely.geometry import LineString, Point


routes_file = "routes.xlsx"
stops_file = "bus-stops.xlsx"
DISTANCE_THRESHOLD_METERS = 50
DEG_TO_M = 111139

# === Загрузка маршрутов ===
xls = pd.ExcelFile(routes_file)
route_lines = []

for sheet in xls.sheet_names:
    df = xls.parse(sheet)
    if {"latitude", "longitude"}.issubset(df.columns):
        df = df.dropna(subset=["latitude", "longitude"])
        df['route_id'] = sheet
        line = LineString(df[["longitude", "latitude"]].values)
        route_lines.append((sheet, line))

# === Загрузка остановок ===
stops_df = pd.read_excel(stops_file)
stops_df = stops_df.dropna(subset=["latitude", "longitude"])

# === Привязка маршрутов к остановкам ===
stop_routes = {}

for route_id, line in route_lines:
    for _, stop in stops_df.iterrows():
        stop_id = stop["bus_stop_id"]
        stop_name = stop["bus_stop_name"]
        stop_point = Point(stop["longitude"], stop["latitude"])
        distance_deg = line.distance(stop_point)
        distance_m = distance_deg * DEG_TO_M
        if distance_m <= DISTANCE_THRESHOLD_METERS:
            if stop_id not in stop_routes:
                stop_routes[stop_id] = {
                    "bus_stop_id": stop_id,
                    "bus_stop_name": stop_name,
                    "routes": set()
                }
            stop_routes[stop_id]["routes"].add(route_id)


records = []
for stop_data in stop_routes.values():
    routes = sorted(list(stop_data["routes"]))
    records.append({
        "Остановка": stop_data["bus_stop_name"],
        "ID остановки": stop_data["bus_stop_id"],
        "Количество маршрутов": len(routes),
        "Маршруты": routes
    })

df_result = pd.DataFrame(records)
df_result.sort_values("Количество маршрутов", ascending=False, inplace=True)

top_hubs = df_result[df_result["Количество маршрутов"] >= df_result["Количество маршрутов"].max()]


df_result.to_excel("пересадочные_узлы_все.xlsx", index=False)
top_hubs.to_excel("топ_пересадочные_узлы.xlsx", index=False)

print("✅ Готово! Сохранено:")
print(" - пересадочные_узлы_все.xlsx")
print(" - топ_пересадочные_узлы.xlsx")
