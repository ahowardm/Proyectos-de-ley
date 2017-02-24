# -*- coding: utf-8 -*-
import requests, datetime
import win32com.client as win32
from bs4 import BeautifulSoup
from selenium import webdriver

proyectos = [9783,10793,9613,10823,8986,9961,10015,10713,11074,10310,11514,10820,9612,8733,6681,10763,8428,7148,10801,9749,9822,9916,9940,9988,10810,9839,10857,11205,11236,10554,10874,11366,10568,10972,11560,11543,11242,11349,11601,11602,11608,11105]
# proyectos = [9783,10793,9613]
estilo = "<style>p.MsoNormal, li.MsoNormal, div.MsoNormal {margin:0cm;margin-bottom:.0001pt;font-size:11.0pt;font-family: \"Calibri\",\"sans-serif\";mso-fareast-language:EN-US;}</style>"
firma = "<p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>Andrés Howard M.</span><b><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#FF5815;mso-fareast-language:ES-CL'>_&nbsp;</span></b><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>Ingeniero de Regulación</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><b><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Gerencia de Regulación y Asuntos Corporativos - Entel</span></b><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Av. Costanera Sur 2760 piso 23, Las Condes</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Tel. +562 2423 2707</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#666666;mso-fareast-language:ES-CL'>Cel. +569 7998 2169</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p><p class=MsoNormal style='line-height:15.0pt;background:white'><span style='font-size:10.0pt;font-family:\"Arial\",\"sans-serif\";color:#0874AF;mso-fareast-language:ES-CL'>ahoward@entel.cl</span><span style='font-size:12.0pt;font-family:\"Arial\",\"sans-serif\";color:#888888;mso-fareast-language:ES-CL'><o:p></o:p></span></p></div>"
mailBody = estilo + "<p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>Estimados, para su sus registros les mando el estado de tramitación de los siguientes Proyectos de Ley.</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>&nbsp</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>&nbsp</span></p>"
autores = ""
for p in proyectos:
    url = "https://www.camara.cl/pley/pley_detalle.aspx?prmID=" + str(p)
    res = requests.get(url)

    if res.status_code == requests.codes.ok:
        proyecto = BeautifulSoup(res.text, "lxml").select("h3.caption")[0].getText().encode("utf-8")
        iniciativa = BeautifulSoup(res.text, "lxml").select("table.tabla > tr.odd > td")[3].getText().encode("utf-8")
        texto = BeautifulSoup(res.text, "lxml").select("#ctl00_mainPlaceHolder_grvtramitacion > tbody > tr")
        temporal = texto[-1].getText(separator = ";").encode("utf-8")

    # mailBody += "<p class=MsoNormal><b><span style=\'font-family:\"Georgia\",\"serif\"\'><a href=\"" + url + "\">" + proyecto + "</a></span></b></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[1] + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[2] + " sesión</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>"+ temporal.split(";")[3] + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[4] + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[5] + "</span></p>"
    mailBody += "<p class=MsoNormal><b><span style=\'font-family:\"Georgia\",\"serif\"\'><a href=\"" + url + "\">" + proyecto + "</a></span></b></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[1] + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>Iniciativa: "+ iniciativa + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[3] + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[4] + "</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>" + temporal.split(";")[5] + "</span></p>"
    mailBody += "<p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>&nbsp</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>&nbsp</span></p>"
mailBody += "<p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>&nbsp</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>Saludos,</span></p><p class=MsoNormal><span style=\'font-family:\"Georgia\",\"serif\"\'>&nbsp</span></p>" + firma

now = datetime.datetime.now()

outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = "ahoward@entel.cl"
mail.Subject = "Proyectos de Ley al %2d" % now.day +  " de %4d%02d" % (now.year, now.month)
mail.HtmlBody = unicode(mailBody, "utf-8")

mail.Display(True)
