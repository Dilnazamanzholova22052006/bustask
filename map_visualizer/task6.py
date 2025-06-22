import pandas as pd
import numpy as np
from shapely.geometry import LineString
from geopy.distance import geodesic
import math

# === Параметры ===
DIST_THRESHOLD_LOOP_METERS = 500     # расстояние между началом и концом маршрута для кольца
ZIGZAG_ANGLE_THRESHOLD = 60          # минимальный угол "излома" (в градусах)
ZIGZAG_COUNT_THRESHOLD = 5           # сколько таких углов считать "зигзагом"

# === Загрузка данных ===
xls = pd.ExcelFile("routes.xlsx")
routes = []

for sheet_name in xls.sheet_names:
    df = xls.parse(sheet_name)
    if {'latitude', 'longitude', 'type'}.issubset(df.columns):
        df['route_id'] = sheet_name
        df = df.dropna(subset=['latitude', 'longitude'])
        routes.append(df)

df_routes = pd.concat(routes, ignore_index=True)

# === Расчёт угла между тремя точками ===
def angle_between(p1, p2, p3):
    a = np.array(p1)
    b = np.array(p2)
    c = np.array(p3)

    ab = a - b
    cb = c - b

    cosine_angle = np.dot(ab, cb) / (np.linalg.norm(ab) * np.linalg.norm(cb))
    angle_rad = np.arccos(np.clip(cosine_angle, -1.0, 1.0))
    return np.degrees(angle_rad)

# === Анализ маршрутов ===
result = []

for (route_id, direction), group in df_routes.groupby(['route_id', 'type']):
    group = group.sort_index()
    coords = list(zip(group['latitude'], group['longitude']))

    if len(coords) < 3:
        route_type = "недостаточно данных"
        comment = "меньше 3 точек"
    else:
        # Проверка на кольцевой маршрут
        start = coords[0]
        end = coords[-1]
        loop_dist = geodesic(start, end).meters

        # Проверка на зигзаги
        sharp_turns = 0
        for i in range(1, len(coords) - 1):
            angle = angle_between(coords[i - 1], coords[i], coords[i + 1])
            if angle < ZIGZAG_ANGLE_THRESHOLD:
                sharp_turns += 1

        if loop_dist <= DIST_THRESHOLD_LOOP_METERS:
            route_type = "кольцевой"
            comment = f"расстояние между началом и концом {int(loop_dist)} м"
        elif sharp_turns >= ZIGZAG_COUNT_THRESHOLD:
            route_type = "зигзагообразный"
            comment = f"найдено {sharp_turns} резких поворотов < {ZIGZAG_ANGLE_THRESHOLD}°"
        else:
            route_type = "линейный"
            comment = f"расстояние между концами {int(loop_dist)} м, поворотов: {sharp_turns}"

    result.append({
        "route_id": route_id,
        "direction": direction,
        "тип_маршрута": route_type,
        "комментарий": comment
    })

df_result = pd.DataFrame(result)
df_result.to_excel("тип_маршрутов.xlsx", index=False)
print("✅ Сохранено: тип_маршрутов.xlsx")
