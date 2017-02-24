# -*- coding: utf-8 -*-
import requests
import win32com.client as win32
from bs4 import BeautifulSoup

url = "https://www.camara.cl/pley/pley_detalle.aspx?prmID=8986"
res = requests.get(url)

if res.status_code == requests.codes.ok:
    proyecto = BeautifulSoup(res.text, "lxml").select("h3.caption")[0].getText().encode("utf-8")
    texto = BeautifulSoup(res.text, "lxml").select("#ctl00_mainPlaceHolder_grvtramitacion > tbody > tr")

temporal = texto[-1].getText(separator = ";").encode("utf-8")
estilo = "<style>p.MsoNormal, li.MsoNormal, div.MsoNormal {margin:0cm;margin-bottom:.0001pt;font-size:11.0pt;font-family: \"Calibri\",\"sans-serif\";mso-fareast-language:EN-US;}</style>"
mailBody = estilo +"<p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>Estimados, para su sus registros les mando el estado de tramitacion de los siguientes Proyectos de Ley.</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>&nbsp</span></p>"
mailBody += "<p class=MsoNormal><b><span style=\'font-family:\"Georgia\",\"serif\"\'><a href=\"" + url + "\">" + proyecto + "</a></span></b></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[1] + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[2] + " sesión</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>"+ temporal.split(";")[3] + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[4] + "</span></p>"
firma = "<p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>Andrés Howard M.</span><b><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#FF5815;mso-fareast-language:ES-CL'>_&nbsp;</span></b><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>Ingeniero de Regulación</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><b><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Gerencia de Regulación y Asuntos Corporativos - Entel</span></b><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Av. Costanera Sur 2760 piso 23, Las Condes</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Tel. +562 2423 2707</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Cel. +569 7998 2169</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>ahoward@entel.cl</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p></div>"
mailBody += "<p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>&nbsp</span></p>" + firma



outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = "ahoward@entel.cl"
mail.Subject = "Proyectos de Ley"
mail.HtmlBody = unicode(mailBody, "utf-8")

mail.Display(True)
