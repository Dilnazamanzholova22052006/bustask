import pandas as pd
from shapely.geometry import LineString, Point
from geopy.distance import geodesic
import warnings


routes_file = "routes.xlsx"
bus_stops_file = "bus-stops.xlsx"

xls = pd.ExcelFile(routes_file)
all_routes = []
for sheet in xls.sheet_names:
    try:
        df = xls.parse(sheet)
        if {'latitude', 'longitude', 'type'}.issubset(df.columns):
            df['route_id'] = sheet
            all_routes.append(df.dropna(subset=['latitude', 'longitude']))
        else:
            print(f" Лист '{sheet}' пропущен — нет нужных колонок")
    except Exception as e:
        print(f"шибка чтения листа {sheet}: {e}")

if not all_routes:
    raise ValueError(" Нет маршрутов с нужными колонками.")

routes_df = pd.concat(all_routes, ignore_index=True)
stops_df = pd.read_excel(bus_stops_file).dropna(subset=["latitude", "longitude"])

# === Найти остановки вблизи маршрутов ===
DISTANCE_THRESHOLD_METERS = 50
matched = []

for (route_id, direction), group in routes_df.groupby(['route_id', 'type']):
    line = LineString(group[['longitude', 'latitude']].values)
    for _, stop in stops_df.iterrows():
        stop_point = Point(stop['longitude'], stop['latitude'])
        distance_deg = line.distance(stop_point)
        approx_meters = distance_deg * 111139
        if approx_meters <= DISTANCE_THRESHOLD_METERS:
            matched.append({
                "route_id": route_id,
                "direction": direction,
                "bus_stop_id": stop["bus_stop_id"],
                "bus_stop_name": stop["bus_stop_name"],
                "latitude": stop["latitude"],
                "longitude": stop["longitude"],
                "distance_m": round(approx_meters, 2)
            })

matched_df = pd.DataFrame(matched)
if matched_df.empty:
    raise ValueError(" Нет совпавших остановок с маршрутами.")

# === Расчёт расстояний между остановками ===
segment_data = []

for (route_id, direction), group in matched_df.groupby(['route_id', 'direction']):
    try:
        line = LineString(
            routes_df[(routes_df['route_id'] == route_id) & (routes_df['type'] == direction)][['longitude', 'latitude']].values
        )
        group = group.copy()
        group['position'] = group.apply(lambda row: line.project(Point(row['longitude'], row['latitude'])), axis=1)
        group = group.sort_values('position').reset_index(drop=True)

        distances = []
        for i in range(len(group) - 1):
            lat1, lon1 = group.loc[i, ['latitude', 'longitude']]
            lat2, lon2 = group.loc[i + 1, ['latitude', 'longitude']]
            dist = geodesic((lat1, lon1), (lat2, lon2)).meters
            distances.append(dist)

        if not distances:
            continue

        segment_data.append({
            "route_id": route_id,
            "direction": direction,
            "start_stop": group.iloc[0]['bus_stop_name'],
            "end_stop": group.iloc[-1]['bus_stop_name'],
            "length_km": round(sum(distances) / 1000, 2),
            "avg_dist_m": round(sum(distances) / len(distances), 1),
            "med_dist_m": round(pd.Series(distances).median(), 1),
        })
    except Exception as e:
        warnings.warn(f" Ошибка в маршруте {route_id} ({direction}): {e}")
        continue

seg_df = pd.DataFrame(segment_data)

# === Переводим в таблицу с колонками по направлениям ===
pivot_df = seg_df.pivot(index='route_id', columns='direction', values=[
    'start_stop', 'end_stop', 'length_km', 'avg_dist_m', 'med_dist_m'
])
pivot_df.columns = ['_'.join(col).strip() for col in pivot_df.columns.values]
pivot_df.reset_index(inplace=True)
pivot_df['full_loop_km'] = pivot_df.get('length_km_forward', 0) + pivot_df.get('length_km_backward', 0)

# === Категоризация по среднему расстоянию в прямом направлении ===
def classify_avg_distance(m):
    if pd.isna(m):
        return 'нет данных'
    elif m <= 300:
        return "до 300 м"
    elif m <= 500:
        return "от 300 до 500 м"
    else:
        return "больше 500 м"

pivot_df['категория_средняя_дистанция'] = pivot_df['avg_dist_m_forward'].apply(classify_avg_distance)

pivot_df.to_excel("среднее_и_медианное_по_остановкам.xlsx", index=False)

# === Группируем маршруты по категориям ===
grouped = pivot_df.groupby('категория_средняя_дистанция').agg(
    перечень_маршрутов=('route_id', list),
    количество_маршрутов=('route_id', 'count')
).reset_index()

grouped.rename(columns={'категория_средняя_дистанция': 'Категория'}, inplace=True)
grouped.to_excel("группировка_по_категориям.xlsx", index=False)

print("✅ Готово! Таблицы сохранены:")
print(" - среднее_и_медианное_по_остановкам.xlsx")
print(" - группировка_по_категориям.xlsx")

