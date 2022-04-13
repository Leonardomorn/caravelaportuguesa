#inicializações
import glob
import json
import lzma
import numpy as np
import os
import pandas as pd
import xlsxwriter

from datetime import datetime
from io import BytesIO
from os import remove
from os.path import exists
from PIL import Image

path = '#caravelaportuguesa/'
url_raw = 'https://raw.githubusercontent.com/heloisafr/caravela_dados/master/caravelaportuguesa/'
url_insta = 'https://www.instagram.com/p/'

#LEITURA DOS JSONS
# cria um dataframe a partir dos dados brutos do .json.xz
# pode demorar 10min

data = []
index = []
files = glob.glob('#caravelaportuguesa/*.json.xz')

c=0
for file in files:

  if file == '#caravelaportuguesa/#caravelaportuguesa.json.xz':
    continue

  conteudo = json.loads(lzma.open(file).read())

  utc = file.replace('.json.xz', '').replace('#caravelaportuguesa/', '')

  timestamp = datetime.strptime(utc, "%Y-%m-%d_%H-%M-%S_UTC")
  typename = conteudo['node']['__typename']
  taken_at = dt_object = datetime.fromtimestamp(conteudo['node']['taken_at_timestamp']) 
  is_video = conteudo['node']['is_video']
  shortcode = conteudo['node']['shortcode']

  text = None
  if len(conteudo['node']['edge_media_to_caption']['edges'])>0:
    if 'node' in conteudo['node']['edge_media_to_caption']['edges'][0]:
      if 'text' in conteudo['node']['edge_media_to_caption']['edges'][0]['node']:
        text = conteudo['node']['edge_media_to_caption']['edges'][0]['node']['text']

  qtde_midias = 1
  lat = ''
  lng = ''
  city = ''
  name = ''
  iphone_struct = False
  if 'iphone_struct' in conteudo['node']:

    iphone_struct = True

    if 'carousel_media_count' in conteudo['node']['iphone_struct']:
      qtde_midias = conteudo['node']['iphone_struct']['carousel_media_count']

    if 'lng' in conteudo['node']['iphone_struct']:
      lng = conteudo['node']['iphone_struct']['lng']
      lat = conteudo['node']['iphone_struct']['lat'] 

    if 'location' in conteudo['node']['iphone_struct']:
      city = conteudo['node']['iphone_struct']['location']['city']
      name = conteudo['node']['iphone_struct']['location']['name']

  name2 = ''
  city2 = ''
  country_code = ''
  address_json = ''
  if 'location' in conteudo['node']:
    if conteudo['node']['location'] is not None:
      name2 = conteudo['node']['location']['name']
      address_json = conteudo['node']['location']['address_json']
      if address_json is not None:
        address_json = json.loads(address_json)
        city2 = address_json['city_name']
        country_code = address_json['country_code']

  # if country_code!='' and country_code!='BR':
  #   continue

  qtde_midias2 = 0
  is_video2 = None
  if 'edge_sidecar_to_children' in conteudo['node']:
    edges = conteudo['node']['edge_sidecar_to_children']['edges']
    qtde_midias2 = len(edges)
    for n in edges:
      is_video2 = n['node']['__typename']

  data.append({'UTC': utc, 'timestamp': timestamp, 'typename': typename, 'taken_at': taken_at,
               'shortcode': shortcode, 'text': text, 'iphone_struct': iphone_struct, 
               'qtde_midias': qtde_midias, 'qtde_midias2': qtde_midias2, 
               'is_video': is_video, 'is_video2': is_video2, 
               'lat': lat, 'lng': lng, 'city': city, 'city2': city2, 
               'country_code': country_code, 'name': name, 'name2': name2})
  index.append(utc) 

  c += 1

  # if c>14:
  #   break
  

df = pd.DataFrame(data, index)
df.sample(1)

# ORDENA O DATAFRAME
df.sort_values(by="UTC", inplace=True)
df.reset_index(drop=True, inplace=True)
df.head()

#GERAÇÃO DA PLANILHA

new_width = 160

def resize(name):

  new_name = 'R_' + name

  if exists(path + new_name):
    return new_name

  im = Image.open(path + name)
  concat = new_width/float(im.size[0])
  size = int((float(im.size[1])*float(concat)))
  resized_im = im.resize((new_width, size), Image.ANTIALIAS)
  new_name = 'R_' + name
  resized_im.save(path + new_name)

  return new_name

workbook = xlsxwriter.Workbook('base_crua.xlsx')
# workbook = xlsxwriter.Workbook('teste.xlsx')
worksheet = workbook.add_worksheet()

header_format = workbook.add_format()
header_format.set_bold()
worksheet.write_row(0, 0, ('TIMESTAMP','LAT','LONG','LOC COUNTRY','LOC CITY','LOC NAME',
                           'GEOLOCATION NAME','LOCATION NAME','STATE','CITY','ROTULO',
                           'NUM','URL INSTA','TEXTO','IMAGENS'), header_format)

# formatação das celular
cell_format = workbook.add_format()
cell_format.set_text_wrap()

col_timestamp = 0
col_lat = 1
col_lng = 2
col_loc_country = 3
col_loc_city = 4
col_loc_name = 5
col_geolocation = 6
col_location = 7
col_state = 8
col_city = 9
col_rotulo = 10
col_numero = 11
col_url = 12
col_texto = 13
col_imagem = 14

worksheet.write_comment(0, col_rotulo, """AVISTAMENTO: é um avistamento na costa brasileira
ACIDENTE: é um acidente na costa brasileira
DUVIDA: quando não há certeza
ALERTA: somente alerta, não é um avistamento ou acidente efetivamente
PORTUGAL: é um avistamento ou acidente, mas não na costa brasileira
REPOST: é um avistamento ou acidente, porém é um REPOST
NADA: não é ocorrência de interesse""", {'width': 300, 'height': 250})

# largura da coluna
worksheet.set_column(col_timestamp, col_timestamp, 25) 
worksheet.set_column(col_rotulo, col_rotulo, 10)
worksheet.set_column(col_loc_city, col_city, 20)
worksheet.set_column(col_texto, col_texto, 50)
worksheet.set_column(col_imagem, col_imagem+10, 25)

row = 1

def write_video(timestamp, texto, col, sequencia=''):
  video = url_raw + timestamp + sequencia + '.mp4'
  worksheet.write(row, col_timestamp, timestamp)
  worksheet.write(row, col_texto, texto, cell_format)
  worksheet.write(row, col_imagem, video)  
  
def write_imagem(timestamp, texto, col, sequencia=''):
  imagem = timestamp + sequencia + '.jpg'
  url = url_raw + imagem
  path_to_img = path + resize(imagem)
  worksheet.write(row, col_timestamp, timestamp)
  worksheet.write(row, col_texto, texto, cell_format)
  worksheet.insert_image(row, col, path_to_img, {'object_position': 1, 'x_offset': 5, 'y_offset': 5, 'url': url})

for linha in df.itertuples():

    # altura da linha
    worksheet.set_row(row, 130)

    timestamp = linha.UTC
    texto = linha.text
    typename = linha.typename
    
    if linha.typename=='GraphVideo':
      write_video(timestamp, texto, col_imagem)
    elif linha.typename=='GraphImage':
      write_imagem(timestamp, texto, col_imagem)
    elif linha.typename=='GraphSidecar':
      for s in range(linha.qtde_midias2):
        coluna = col_imagem + s
        sequencia = '_' + str(s + 1)
        if linha.typename=='GraphVideo':
          write_video(timestamp, texto, coluna, sequencia)
        else:
          write_imagem(timestamp, texto, coluna, sequencia)
    else:
      raise Exception('ops typename vazio' + str(timestamp))    
      
    worksheet.write(row, col_lat, linha.lat)
    worksheet.write(row, col_lng, linha.lng)
    worksheet.write(row, col_loc_country, linha.country_code)
    worksheet.write(row, col_loc_city, linha.city if linha.city!='' else linha.city2, cell_format)
    worksheet.write(row, col_loc_name, linha.name if linha.name!='' else linha.name2, cell_format)
    worksheet.write(row, col_url, url_insta + linha.shortcode)

    row += 1

    # if row>5:
    #   break

    # try:
    # except Exception as err:
    #   print('erro:', utc)
    #   print(f"Unexpected {err}, {type(err)}")
    #   print("")


source = ['AVISTAMENTO','ACIDENTE', 'ALERTA', 'PORTUGAL', 'REPOST', 'NADA']
worksheet.data_validation(1, col_rotulo, row, col_rotulo, {'validate': 'list', 'source': source})

workbook.close()