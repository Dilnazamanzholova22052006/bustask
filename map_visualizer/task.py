import pandas as pd
import folium
from folium.plugins import MarkerCluster
from geopy.distance import geodesic

# === Загружаем ВСЕ ЛИСТЫ из маршрутов ===
xls = pd.ExcelFile("routes.xlsx")
all_routes = []

for sheet_name in xls.sheet_names:
    df = xls.parse(sheet_name)
    df['route_id'] = str(sheet_name).strip()
    all_routes.append(df)

routes_df = pd.concat(all_routes, ignore_index=True)

# === Загружаем остановки ===
stops_df = pd.read_excel("bus-stops.xlsx")

# === Очищаем данные ===
routes_clean = routes_df.dropna(subset=["latitude", "longitude"])
stops_clean = stops_df.dropna(subset=["latitude", "longitude"])

# === Центр карты ===
center_lat = routes_clean['latitude'].mean()
center_lon = routes_clean['longitude'].mean()
m = folium.Map(location=[center_lat, center_lon], zoom_start=12)

# === Остановки на карте ===
marker_cluster = MarkerCluster().add_to(m)
for _, row in stops_clean.iterrows():
    folium.Marker(
        location=[row['latitude'], row['longitude']],
        popup=row.get('bus_stop_name', 'Остановка'),
        icon=folium.Icon(color='blue', icon='info-sign')
    ).add_to(marker_cluster)

# === Маршруты и их длины ===
results = []

for (route_id, direction), group in routes_clean.groupby(['route_id', 'type']):
    group_sorted = group.reset_index(drop=True)
    total_distance_km = 0.0
    points = []

    for i in range(len(group_sorted) - 1):
        lat1, lon1 = group_sorted.loc[i, ['latitude', 'longitude']]
        lat2, lon2 = group_sorted.loc[i + 1, ['latitude', 'longitude']]
        dist = geodesic((lat1, lon1), (lat2, lon2)).kilometers
        total_distance_km += dist
        points.append([lat1, lon1])
    if len(group_sorted) >= 2:
        points.append([lat2, lon2])  # добавим последнюю точку

        # рисуем маршрут на карту
        folium.PolyLine(
            points,
            color='red' if direction == 'forward' else 'green',
            weight=3,
            opacity=0.7,
            popup=f"{route_id} ({direction})\nДлина: {round(total_distance_km, 2)} км"
        ).add_to(m)

        # сохраняем в таблицу
        results.append({
            "route_id": route_id,
            "direction": direction,
            "length_km": round(total_distance_km, 2)
        })

# === Сохраняем карту и таблицу ===
m.save("bus_routes_map.html")
print("✅ Карта сохранена в bus_routes_map.html")

length_df = pd.DataFrame(results)
length_df.to_excel("маршруты_с_длиной.xlsx", index=False)
print("✅ Таблица сохранена в маршруты_с_длиной.xlsx")
