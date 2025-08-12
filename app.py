import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import math
import io
from typing import List, Tuple
import re
from io import BytesIO

def haversine_distance(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """
    Menghitung jarak antara dua titik koordinat menggunakan rumus Haversine
    Returns: jarak dalam meter
    """
    R = 6371000  # Radius bumi dalam meter
    
    # Konversi derajat ke radian
    lat1_rad = math.radians(lat1)
    lon1_rad = math.radians(lon1)
    lat2_rad = math.radians(lat2)
    lon2_rad = math.radians(lon2)
    
    # Perbedaan koordinat
    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad
    
    # Rumus Haversine
    a = math.sin(dlat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon/2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
    
    return R * c

def parse_kml_coordinates(kml_content: str) -> List[Tuple[float, float]]:
    """
    Mengekstrak koordinat dari file KML
    """
    try:
        root = ET.fromstring(kml_content)
        
        # Namespace KML
        namespace = {'kml': 'http://www.opengis.net/kml/2.2'}
        
        coordinates = []
        
        # Cari semua elemen coordinates
        for coord_elem in root.findall('.//kml:coordinates', namespace):
            coord_text = coord_elem.text.strip()
            
            # Split berdasarkan whitespace dan koma
            coord_parts = re.split(r'[\s,]+', coord_text)
            
            # Parse koordinat (format: longitude,latitude,altitude)
            for i in range(0, len(coord_parts)-2, 3):
                try:
                    lon = float(coord_parts[i])
                    lat = float(coord_parts[i+1])
                    coordinates.append((lat, lon))
                except (ValueError, IndexError):
                    continue
        
        return coordinates
    except Exception as e:
        st.error(f"Error parsing KML: {str(e)}")
        return []

def kml_to_csv(kml_content: str) -> pd.DataFrame:
    """
    Konversi KML ke CSV dengan perhitungan jarak dan titik tengah
    """
    coordinates = parse_kml_coordinates(kml_content)
    
    if not coordinates:
        return pd.DataFrame()
    
    data = []
    
    for i in range(len(coordinates)):
        lat, lon = coordinates[i]
        
        row = {
            'latitude': lat,
            'longitude': lon,
            'length': None,
            'latitude mid': None,
            'longitude mid': None
        }
        
        # Hitung jarak dan titik tengah jika bukan baris terakhir
        if i < len(coordinates) - 1:
            next_lat, next_lon = coordinates[i + 1]
            
            # Hitung jarak menggunakan Haversine
            distance = haversine_distance(lat, lon, next_lat, next_lon)
            row['length'] = round(distance)
            
            # Hitung titik tengah
            row['latitude mid'] = (lat + next_lat) / 2
            row['longitude mid'] = (lon + next_lon) / 2
        
        data.append(row)
    
    return pd.DataFrame(data)

def csv_to_kml(df: pd.DataFrame) -> str:
    """
    Konversi CSV ke KML
    """
    kml_template = '''<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">
<Document>
	<name>length.kml</name>
	<StyleMap id="m_ylw-pushpin">
		<Pair>
			<key>normal</key>
			<styleUrl>#s_ylw-pushpin0</styleUrl>
		</Pair>
		<Pair>
			<key>highlight</key>
			<styleUrl>#s_ylw-pushpin_hl</styleUrl>
		</Pair>
	</StyleMap>
	<StyleMap id="m_ylw-pushpin0">
		<Pair>
			<key>normal</key>
			<styleUrl>#s_ylw-pushpin</styleUrl>
		</Pair>
		<Pair>
			<key>highlight</key>
			<styleUrl>#s_ylw-pushpin_hl0</styleUrl>
		</Pair>
	</StyleMap>
	<Style id="s_ylw-pushpin">
		<IconStyle>
			<Icon>
			</Icon>
		</IconStyle>
		<ListStyle>
		</ListStyle>
	</Style>
	<Style id="s_ylw-pushpin0">
		<IconStyle>
			<Icon>
			</Icon>
		</IconStyle>
		<ListStyle>
		</ListStyle>
	</Style>
	<Style id="s_ylw-pushpin_hl">
		<IconStyle>
			<Icon>
			</Icon>
		</IconStyle>
		<ListStyle>
		</ListStyle>
	</Style>
	<Style id="s_ylw-pushpin_hl0">
		<IconStyle>
			<Icon>
			</Icon>
		</IconStyle>
		<ListStyle>
		</ListStyle>
	</Style>
	<Folder>
		<name>length</name>
{placemarks}
		<atom:link rel="app" href="https://www.google.com/earth/about/versions/#earth-pro" title="Google Earth Pro 7.3.6.10201"></atom:link>
	</Folder>
</Document>
</kml>'''
    
    placemarks = []
    
    for index, row in df.iterrows():
        if pd.notna(row['length']) and pd.notna(row['latitude']) and pd.notna(row['longitude']):
            style_ref = "#m_ylw-pushpin0" if index == 0 else "#m_ylw-pushpin"
            
            placemark = f'''		<Placemark>
			<name>{int(row['length'])}</name>
			<LookAt>
				<longitude>{row['longitude']}</longitude>
				<latitude>{row['latitude']}</latitude>
				<altitude>0</altitude>
				<heading>-1.521859235210106e-05</heading>
				<tilt>14.4237821896138</tilt>
				<range>206.1914177482452</range>
				<gx:altitudeMode>relativeToSeaFloor</gx:altitudeMode>
			</LookAt>
			<styleUrl>{style_ref}</styleUrl>
			<Point>
				<gx:drawOrder>1</gx:drawOrder>
				<coordinates>{row['longitude']},{row['latitude']},0</coordinates>
			</Point>
		</Placemark>'''
            
            placemarks.append(placemark)
    
    return kml_template.format(placemarks='\n'.join(placemarks))

def main():
    st.set_page_config(
        page_title="MappingPro",
        page_icon="🗺️",
        layout="wide"
    )
    
    st.title("🗺️ MappingPro")
    st.markdown("**Aplikasi Konversi KML ↔ CSV untuk Data Geospasial** - Created by [Rendi Gunawan](https://github.com/rendinawan)")
    
    # Sidebar untuk navigasi
    st.sidebar.title("Tools")
    tool_choice = st.sidebar.selectbox(
        "Pilih Tool:",
        ["KML ke CSV/Excel", "CSV/Excel ke KML"]
    )
    
    if tool_choice == "KML ke CSV/Excel":
        st.header("📁 Konversi KML ke CSV/Excel")
        st.markdown("Upload file KML untuk dikonversi ke format CSV atau Excel dengan perhitungan jarak dan titik tengah.")
        
        uploaded_file = st.file_uploader(
            "Pilih file KML", 
            type=['kml'],
            help="Upload file KML yang berisi data koordinat"
        )
        
        # Pilihan format output
        col1, col2 = st.columns(2)
        with col1:
            output_format = st.radio(
                "Format Output:",
                ["CSV", "Excel (.xlsx)"],
                horizontal=True
            )
        
        if uploaded_file is not None:
            try:
                # Baca konten file
                kml_content = uploaded_file.getvalue().decode('utf-8')
                
                # Konversi ke CSV
                df = kml_to_csv(kml_content)
                
                if not df.empty:
                    st.success("✅ File KML berhasil dikonversi!")
                    
                    # Tampilkan preview data
                    st.subheader("Preview Data:")
                    st.dataframe(df, use_container_width=True)
                    
                    # Informasi statistik
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Titik", len(df))
                    with col2:
                        total_distance = df['length'].sum()
                        if pd.notna(total_distance):
                            st.metric("Total Jarak", f"{total_distance:.1f} m")
                    with col3:
                        avg_distance = df['length'].mean()
                        if pd.notna(avg_distance):
                            st.metric("Rata-rata Jarak", f"{avg_distance:.1f} m")
                    
                    # Download berdasarkan format yang dipilih
                    if output_format == "CSV":
                        csv_buffer = io.StringIO()
                        df.to_csv(csv_buffer, sep=';', index=False)
                        csv_string = csv_buffer.getvalue()
                        
                        st.download_button(
                            label="📥 Download CSV",
                            data=csv_string,
                            file_name="route.csv",
                            mime="text/csv"
                        )
                    else:  # Excel
                        excel_buffer = BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            df.to_excel(writer, sheet_name='Route Data', index=False)
                        excel_data = excel_buffer.getvalue()
                        
                        st.download_button(
                            label="📥 Download Excel",
                            data=excel_data,
                            file_name="route.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                else:
                    st.error("❌ Tidak dapat mengekstrak koordinat dari file KML")
                    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    
    elif tool_choice == "CSV/Excel ke KML":
        st.header("📁 Konversi CSV/Excel ke KML")
        st.markdown("Upload file CSV atau Excel untuk dikonversi ke format KML.")
        
        uploaded_file = st.file_uploader(
            "Pilih file CSV atau Excel", 
            type=['csv', 'xlsx', 'xls'],
            help="Upload file CSV atau Excel dengan kolom: length, latitude, longitude"
        )
        
        if uploaded_file is not None:
            try:
                # Tentukan jenis file dan baca data
                file_extension = uploaded_file.name.split('.')[-1].lower()
                
                if file_extension == 'csv':
                    # Baca CSV dengan berbagai separator
                    try:
                        df = pd.read_csv(uploaded_file, sep=';')
                    except:
                        uploaded_file.seek(0)  # Reset file pointer
                        df = pd.read_csv(uploaded_file, sep=',')
                    st.info("📄 File CSV berhasil dibaca")
                    
                elif file_extension in ['xlsx', 'xls']:
                    # Baca Excel file
                    excel_file = pd.ExcelFile(uploaded_file)
                    
                    # Jika ada multiple sheets, biarkan user memilih
                    if len(excel_file.sheet_names) > 1:
                        selected_sheet = st.selectbox(
                            "Pilih Sheet:",
                            excel_file.sheet_names
                        )
                        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
                    else:
                        df = pd.read_excel(uploaded_file, sheet_name=0)
                    
                    st.info(f"📊 File Excel berhasil dibaca ({len(excel_file.sheet_names)} sheet tersedia)")
                
                st.success("✅ File berhasil dibaca!")
                
                # Tampilkan preview data
                st.subheader("Preview Data:")
                st.dataframe(df, use_container_width=True)
                
                # Validasi kolom yang diperlukan
                required_columns = ['length', 'latitude', 'longitude']
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    st.error(f"❌ Kolom berikut tidak ditemukan: {', '.join(missing_columns)}")
                    st.info("💡 File harus memiliki kolom: length, latitude, longitude")
                    
                    # Tampilkan kolom yang tersedia
                    st.subheader("Kolom yang tersedia:")
                    st.write(list(df.columns))
                else:
                    # Konversi ke KML
                    kml_content = csv_to_kml(df)
                    
                    # Informasi statistik
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total Points", len(df))
                    with col2:
                        valid_points = df.dropna(subset=['length', 'latitude', 'longitude'])
                        st.metric("Valid Points", len(valid_points))
                    
                    # Download KML
                    st.download_button(
                        label="📥 Download KML",
                        data=kml_content,
                        file_name="length.kml",
                        mime="application/vnd.google-earth.kml+xml"
                    )
                    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.info("💡 Pastikan file Excel tidak dalam keadaan terbuka di aplikasi lain")
    
    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: gray;'>
            <p>MappingPro v1.1 - Tool Konversi Data Geospasial</p>
            <p>Mendukung konversi KML ↔ CSV/Excel dengan perhitungan Haversine</p>
        </div>
        """, 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
