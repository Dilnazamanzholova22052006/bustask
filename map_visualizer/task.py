import pandas as pd
import folium
from folium.plugins import MarkerCluster


routes_df = pd.read_excel("routes.xlsx")
stops_df = pd.read_excel("bus-stops.xlsx")


routes_clean = routes_df.dropna(subset=["latitude", "longitude"])
stops_clean = stops_df.dropna(subset=["latitude", "longitude"])


center_lat = routes_clean['latitude'].mean()
center_lon = routes_clean['longitude'].mean()


m = folium.Map(location=[center_lat, center_lon], zoom_start=12)


marker_cluster = MarkerCluster().add_to(m)
for _, row in stops_clean.iterrows():
    folium.Marker(
        location=[row['latitude'], row['longitude']],
        popup=row.get('bus_stop_name', 'Остановка'),
        icon=folium.Icon(color='blue', icon='info-sign')
    ).add_to(marker_cluster)


for (route_id, direction), group in routes_clean.groupby(['route_id', 'type']):
    points = group[['latitude', 'longitude']].values.tolist()
    folium.PolyLine(
        points,
        color='red' if direction == 'forward' else 'green',
        weight=3,
        opacity=0.7,
        popup=f"{route_id} ({direction})"
    ).add_to(m)


m.save("bus_routes_map.html")
