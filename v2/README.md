# Diese Datei ist von Bazarbekov Akhmet vorbereitet
# Diese API ist dafür designiert, um Dateien von DOC-V nach Excel zu verwandeln

# Erstellungsdatum: 1 Jänner, 2024

# Erste Schritte zum Umsetzen:

# Verbindung zum Datentrâger-Server (über SSH)
# Ziehen nach Documents/api/excel
# Virtuelle Umgebung aktivieren: source myenv/bin/activate
# WSGI Server mit folgenden Parametern einstellen: gunicorn -w 4 -b host:port app:app
# HOST und PORT sind in Flask-Application definiert

# Optional:
#   -logging - schaltet logging ein
#   -limit - verteidigt jemal den Server von DDoS (Distributed Denial of Service) Attacks

# gunicorn deaktivieren: pkill gunicorn

# Beschreibungen der Funktionen:
#
# 1. set_border - setzt alle Seiten der Zelle/Zellen auf schmale schwarze Linie
# 2. format_row - Formatiert die sogenannte Stroke besprechend zum Format von LOTUS REESTR (von Stawitzkaja E. gesendet)
# 3. hide_sheets - hindert die ausgewählte Excel-Papier
# 4. 

# API updaten: der Muster fur das Hineinladen der PY Dateien an den Server lautet: scp -r ~/Documents/api_excel/template_outer.xlsx radmin@192.168.30.19:~/Documents/api_excel