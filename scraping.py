# -*- coding: utf-8 -*-
import requests, datetime, re
import win32com.client as win32
from bs4 import BeautifulSoup

now = datetime.datetime.now()
### Normas Generales ###
url = "http://www.doe.cl/sumario_por_seccion.php?dia=" + "%02d" % now.day + "&mes=" + "%02d" % now.month + "&anio=2017&seccion=1"
res = requests.get(url)

if res.status_code == requests.codes.ok:
    texto = BeautifulSoup(res.text, "lxml").find_all("td", { "class": re.compile(r"^title[4-5]|referencia$")})

    mailBody = "<p>Estimados, el día de hoy se han registrado las siguientes publicaciones en el Diario Oficial.</p><style>p.MsoNormal, li.MsoNormal, div.MsoNormal{margin:0cm;margin-bottom:.0001pt;font-size:11.0pt;font-family:""Calibri"",""sans-serif"";mso-fareast-language:EN-US;}a:link, span.MsoHyperlink{mso-style-priority:99;color:blue;text-decoration:underline;}</style>"
    mailBody += "<h2>Normas Generales</h2><br><a href=\"http://www.doe.cl/sumario_por_seccion.php?dia=" + "%02d" % now.day + "&mes=" + "%02d" % now.month + "&anio=2017&seccion=1\"> Normas Generales</a>"
    print "Normas Generales"
    for i in range(len(texto)):
        print texto[i].getText().encode("utf-8")
        mailBody += "<p>" + texto[i].getText().encode("utf-8") + "</p>"

### Normas Particulares ###
url = "http://www.doe.cl/sumario_por_seccion.php?dia=" + "%02d" % now.day + "&mes=" + "%02d" % now.month + "&anio=2017&seccion=2"
res = requests.get(url)

if res.status_code == requests.codes.ok:
    texto = BeautifulSoup(res.text, "lxml").find_all("td", { "class": re.compile(r"^title[4-5]|referencia$")})

    mailBody += "<h2>Normas Particulares</h2><br><a href=\"http://www.doe.cl/sumario_por_seccion.php?dia=" + "%02d" % now.day + "&mes=" + "%02d" % now.month + "&anio=2017&seccion=2\"> Normas Particulares</a>"
    print "Normas Particulares"
    for i in range(len(texto)):
        print texto[i].getText().encode("utf-8")
        mailBody += "<p>" + texto[i].getText().encode("utf-8") + "</p>"

    mailBody += "<p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>Andrés Howard M.</span><b><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#FF5815;mso-fareast-language:ES-CL'>_&nbsp;</span></b><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>Ingeniero de Regulación</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><b><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Gerencia de Regulación y Asuntos Corporativos - Entel</span></b><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Av. Costanera Sur 2760 piso 23, Las Condes</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Tel. +562 2423 2707</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Cel. +569 7998 2169</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>ahoward@entel.cl</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p></div>"

outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = "ahoward@entel.cl"
mail.Subject = "Diario Oficial"
mail.HtmlBody = unicode(mailBody, "utf-8")

mail.Display(True)
