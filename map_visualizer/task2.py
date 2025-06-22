import pandas as pd
from shapely.geometry import LineString, Point
from geopy.distance import geodesic

xls = pd.ExcelFile("routes.xlsx")
routes_list = []

for sheet_name in xls.sheet_names:
    df = xls.parse(sheet_name)
    df['route_id'] = str(sheet_name).strip()
    if 'type' not in df.columns:
        df['type'] = 'forward'
    routes_list.append(df)

routes_df = pd.concat(routes_list, ignore_index=True)
routes_df = routes_df.dropna(subset=["latitude", "longitude"])


bus_stops_df = pd.read_excel("bus-stops.xlsx")
bus_stops_df = bus_stops_df.dropna(subset=["latitude", "longitude"])

DISTANCE_THRESHOLD_METERS = 50


matched_stops_geo = []

for (route_id, direction), group in routes_df.groupby(['route_id', 'type']):
    line = LineString(group[['longitude', 'latitude']].values)
    
    for _, stop in bus_stops_df.iterrows():
        stop_point = Point(stop['longitude'], stop['latitude'])
        distance_deg = line.distance(stop_point)
        approx_meters = distance_deg * 111139

        if approx_meters <= DISTANCE_THRESHOLD_METERS:
            matched_stops_geo.append({
                "route_id": route_id,
                "direction": direction,
                "bus_stop_id": stop["bus_stop_id"],
                "bus_stop_name": stop["bus_stop_name"],
                "latitude": stop["latitude"],
                "longitude": stop["longitude"],
                "distance_m": round(approx_meters, 2)
            })

matched_df = pd.DataFrame(matched_stops_geo)

# === Сортировка остановок по линии и расчёт расстояний между ними
segment_summary = []

for (route_id, direction), group in matched_df.groupby(['route_id', 'direction']):
    line = LineString(routes_df[(routes_df['route_id'] == route_id) & (routes_df['type'] == direction)][['longitude', 'latitude']].values)

    group = group.copy()
    group['position_on_line'] = group.apply(
        lambda row: line.project(Point(row['longitude'], row['latitude'])),
        axis=1
    )
    group_sorted = group.sort_values(by='position_on_line').reset_index(drop=True)

    segment_distance = 0.0
    for i in range(len(group_sorted) - 1):
        lat1, lon1 = group_sorted.loc[i, ['latitude', 'longitude']]
        lat2, lon2 = group_sorted.loc[i + 1, ['latitude', 'longitude']]
        segment_distance += geodesic((lat1, lon1), (lat2, lon2)).kilometers

    segment_summary.append({
        "route_id": route_id,
        "direction": direction,
        "start_stop": group_sorted.iloc[0]["bus_stop_name"],
        "end_stop": group_sorted.iloc[-1]["bus_stop_name"],
        "length_km": round(segment_distance, 2)
    })


segment_df = pd.DataFrame(segment_summary)

final_df = segment_df.pivot(index='route_id', columns='direction', values=['start_stop', 'end_stop', 'length_km'])
final_df.columns = ['_'.join(col).strip() for col in final_df.columns.values]
final_df.reset_index(inplace=True)
final_df['full_loop_km'] = final_df.get('length_km_forward', 0) + final_df.get('length_km_backward', 0)

final_df.to_excel("длины_всех_маршрутов.xlsx", index=False)
print("✅ Сохранено: длины_всех_маршрутов.xlsx")
# === Функция классификации маршрутов по длине
def classify_length(length):
    if length < 10:
        return "менее 10 км"
    elif length < 25:
        return "10–25 км"
    elif length < 35:
        return "25–35 км"
    elif length < 50:
        return "35–50 км"
    else:
        return "более 50 км"

# === Применяем классификацию
final_df['Диапазон'] = final_df['full_loop_km'].apply(classify_length)

# === Группируем по диапазонам
classification_df = final_df.groupby('Диапазон').agg({
    'route_id': lambda x: list(x),
    'full_loop_km': 'count'
}).reset_index()

classification_df.columns = ['Диапазон', 'Перечень маршрутов', 'Количество маршрутов']


classification_df.to_excel("классификация_маршрутов.xlsx", index=False)
print("✅ Сохранено: классификация_маршрутов.xlsx")
