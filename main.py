from oauth2client.service_account import ServiceAccountCredentials
import gspread
import json
import time

scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

credentials = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scopes) # Credentials key from Google Sheets file
file = gspread.authorize(credentials) # Authenticate the JSON key with gspread
sheet = file.open("Engenharia de Software - Desafio João Vitor Carvalho Domingos") #open sheet
sheet = sheet.sheet1
print("Calculando a média da provas de cada estudante para arquivo de planilha. Pode demorar alguns minutos. Por favor, seja paciente. Para acompanhar o progresso, abra essa planilha: https://docs.google.com/spreadsheets/d/1Amr1VKaAgwn0NMXaVNHMrepKyMOwH6QT__AWKYMGbkk/edit#gid=0")
for students in range(4, 28):
   time.sleep(5)
   media = ( int(sheet.acell("D{}".format(students)).value) * ((int(sheet.acell("C{}".format(students)).value) - 1 ) / 3) + int(sheet.acell("E{}".format(students)).value) * ((int(sheet.acell("C{}".format(students)).value) - 1 ) / 3) + int(sheet.acell("F{}".format(students)).value) * ((int(sheet.acell("C{}".format(students)).value) - 1 ) / 3) ) / 3
   if (int(sheet.acell("C{}".format(students)).value) >= 15):
      sheet.update_acell('G{}'.format(students), 'Reprovado por Falta')
      sheet.update_acell('H{}'.format(students), 0)
      sheet.update_acell('I{}'.format(students), "{:.0f}".format(media))
      studentsrow = sheet.row_values(students)
   else:
      time.sleep(5)
      if (int(sheet.acell("C{}".format(students)).value) >= 15):
         time.sleep(5)
         sheet.update_acell('G{}'.format(students), 'Reprovado por Falta')
         sheet.update_acell('H{}'.format(students), 0)
         sheet.update_acell('I{}'.format(students), "{:.0f}".format(media))
         studentsrow = sheet.row_values(students)
      elif media < 5:
         time.sleep(5)
         sheet.update_acell('G{}'.format(students), 'Reprovado por Nota')
         sheet.update_acell('H{}'.format(students), 0)
         sheet.update_acell('I{}'.format(students), "{:.0f}".format(media))
         studentsrow = sheet.row_values(students)
      elif media >= 5 and media < 7:
         time.sleep(5)
         naf = 2 * 5 - media
         sheet.update_acell('G{}'.format(students), 'Exame Final')
         sheet.update_acell('H{}'.format(students), naf)
         sheet.update_acell('I{}'.format(students), "{:.0f}".format(media))
         studentsrow = sheet.row_values(students)
      elif media >= 7:
         time.sleep(5)
         sheet.update_acell('G{}'.format(students), 'Aprovado')
         sheet.update_acell('H{}'.format(students), 0)
         sheet.update_acell('I{}'.format(students), "{:.0f}".format(media))
         studentsrow = sheet.row_values(students)
print("Pronto!")
