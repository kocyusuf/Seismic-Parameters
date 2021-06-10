import webbrowser
import folium
from folium.plugins import Draw, MousePosition
import io

def indexMap():
    m = folium.Map(location=[39, 36], zoom_start=6)
    folium.LatLngPopup().add_to(m)
    draw = Draw(
        draw_options={
            'polyline': False,
            'rectangle': False,
            'polygon': False,
            'circle': False,
            'marker': True,
            'circlemarker': False},
        edit_options={'edit': False})
    m.add_child(draw)
    data = io.BytesIO()
    m.save(data, close_file=False)
    formatter = "function(num) {return L.Util.formatNum(num, 3) + ' º ';};"
    MousePosition(
        position="topright",
        separator=" | ",
        empty_string="NaN",
        lng_first=True,
        num_digits=20,
        prefix="Coordinates:",
        lat_formatter=formatter,
        lng_formatter=formatter,
    ).add_to(m)
    m.save("trindex.html")
