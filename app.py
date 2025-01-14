from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import time
import os
from src.functions import readData

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
  os.makedirs(app.config['UPLOAD_FOLDER'])


def delete_old_files(directory, extension, age_minutes=45):
  current_time = time.time()
  age_seconds = age_minutes * 60

  for filename in os.listdir(directory):
    file_path = os.path.join(directory, filename)
    if os.path.isfile(file_path):
      file_age = current_time - os.path.getmtime(file_path)
      if file_age > age_seconds:
          os.remove(file_path)
          print(f"Deleted {file_path}")

@app.route('/')
def home():
  return render_template('/home/home.html')

@app.route('/data-catastro')
def data_catastro():
  return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
  delete_old_files(app.config['UPLOAD_FOLDER'], '.xlsx', 45)
  if 'file_demanda' not in request.files or 'file_pdr' not in request.files:
      return jsonify({'error': 'No file part'}), 400
  file_demanda = request.files['file_demanda']
  file_pdr = request.files['file_pdr']
  if file_demanda.filename == '' or file_pdr.filename == '':
      return jsonify({'error': 'No selected file'}), 400
  filenames = []
  for file in [file_demanda, file_pdr]:
    if file:
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        filenames.append(filename)

  file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_demanda.filename)
  file_path_pdr = os.path.join(app.config['UPLOAD_FOLDER'], file_pdr.filename)
  list_sectors = readData.get_all_sectors(file_path)
  list_columns_demanda = readData.get_all_columns(file_path)
  list_columns_postes = readData.get_all_columns(file_path_pdr)
  list_columns_tanquillas = readData.get_all_columns(file_path_pdr, 1)

  return jsonify({'filenames': filenames, 'list_sectors' : list_sectors, 'list_columns_demanda': list_columns_demanda, 'list_columns_postes': list_columns_postes, 'list_columns_tanquillas': list_columns_tanquillas}), 200

@app.route('/process', methods=['POST'])
def process_file():
  name_sheet_edificaciones = "EDIFICACIONES"
  list_column_origin = ['cod_parc', 'parroquia', 'Tipo de Vía', 'Dirección', 'Uso Inmueble', 'Nombre Inmueble', 'Tipo Inmueble', 'Fecha', 'Total Unidades', 'Total Residenciales', 
  'Total Comerciales', 'Total Oficinas', 'Total Medico Asistencial', 'Total Gubernamental', 'Total Educativo', 'Manzana asignada', 'Total Religioso', 'Total construccion', 
  'Total abandonadas', 'Total Plaza', 'Total Parques', 'sectores']

  list_column_origin_postes = ['Punto de referencia', 'Tipo de Poste', 'El poste tiene código', 'Fecha Levantaamiento', 'Nro. Poste', 'Manzana Asignada', 'sector']
  sheet_postes_origin = 'Postes_0'


  sheet_tanquilla_origin = "Tanquillas_1"
  list_column_origin_tanquilla = ['Punto de Referencia', 'Tipo de tanquilla',	'Fecha Levantamiento', 'sector',	'Manzana asignada']

  data = request.get_json()
  filenames = data.get('filenames')
  area = data.get('area')
  if not filenames or len(filenames) != 2:
    return jsonify({'error': 'No filename provided'}), 400

  name_file = area + '_catastro.xlsx'
  input_paths = [os.path.join(app.config['UPLOAD_FOLDER'], filename) for filename in filenames]
  output_path = os.path.join(app.config['UPLOAD_FOLDER'], name_file)

  for input_path in input_paths:
    if not os.path.exists(input_path):
      return jsonify({'error': f'File {input_path} not found'}), 404

  
  readData.create_one_exel_area(area, input_paths[0], input_paths[1], list_column_origin, name_sheet_edificaciones,
    sheet_postes_origin, sheet_tanquilla_origin, list_column_origin_postes, list_column_origin_tanquilla,
    name_file)
  
  return jsonify({'processed_file': area + '_catastro.xlsx'}), 200


@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({'error': 'File not found'}), 404

if __name__ == '__main__':
  app.run(host='localhost', port=3000, debug=True)
 