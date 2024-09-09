import pyautogui as pag
import keyboard, time, requests, subprocess, ctypes, smtplib, shutil, pyttsx3, threading, imageio, os, sys, pynput, pyperclip
from ctypes import wintypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter import *
from tkinter import messagebox
from PIL import Image, ImageTk, ImageSequence


def obter_usuario():
    return os.getenv('USERNAME')


usuario_atual = obter_usuario()

URL_LED = '<ip_led>'
URL_WPP = '<api_whatsapp>'
email_remetente = "<email>"
senha = '<pass>'
email_destinatario_gravacao = ["<email1>",
                      "<email2>",
]
email_destinatario_programa = "<email>"
assunto = "ZOOM"
mensagem_inicio = """\
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html dir="ltr" xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="pt">
 <head>
  <meta charset="UTF-8">
  <meta content="width=device-width, initial-scale=1" name="viewport">
  <meta name="x-apple-disable-message-reformatting">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta content="telephone=no" name="format-detection">
  <title>Novo modelo</title><!--[if (mso 16)]>
    <style type="text/css">
    a {text-decoration: none;}
    </style>
    <![endif]--><!--[if gte mso 9]><style>sup { font-size: 100% !important; }</style><![endif]--><!--[if gte mso 9]>
<xml>
    <o:OfficeDocumentSettings>
    <o:AllowPNG></o:AllowPNG>
    <o:PixelsPerInch>96</o:PixelsPerInch>
    </o:OfficeDocumentSettings>
</xml>
<![endif]-->
  <style type="text/css">
#outlook a {
	padding:0;
}
.es-button {
	mso-style-priority:100!important;
	text-decoration:none!important;
}
a[x-apple-data-detectors] {
	color:inherit!important;
	text-decoration:none!important;
	font-size:inherit!important;
	font-family:inherit!important;
	font-weight:inherit!important;
	line-height:inherit!important;
}
.es-desk-hidden {
	display:none;
	float:left;
	overflow:hidden;
	width:0;
	max-height:0;
	line-height:0;
	mso-hide:all;
}
@media only screen and (max-width:600px) {p, ul li, ol li, a { line-height:150%!important } h1, h2, h3, h1 a, h2 a, h3 a { line-height:120%!important } h1 { font-size:36px!important; text-align:left } h2 { font-size:26px!important; text-align:left } h3 { font-size:20px!important; text-align:left } .es-header-body h1 a, .es-content-body h1 a, .es-footer-body h1 a { font-size:36px!important; text-align:left } .es-header-body h2 a, .es-content-body h2 a, .es-footer-body h2 a { font-size:26px!important; text-align:left } .es-header-body h3 a, .es-content-body h3 a, .es-footer-body h3 a { font-size:20px!important; text-align:left } .es-menu td a { font-size:12px!important } .es-header-body p, .es-header-body ul li, .es-header-body ol li, .es-header-body a { font-size:14px!important } .es-content-body p, .es-content-body ul li, .es-content-body ol li, .es-content-body a { font-size:14px!important } .es-footer-body p, .es-footer-body ul li, .es-footer-body ol li, .es-footer-body a { font-size:14px!important } .es-infoblock p, .es-infoblock ul li, .es-infoblock ol li, .es-infoblock a { font-size:12px!important } *[class="gmail-fix"] { display:none!important } .es-m-txt-c, .es-m-txt-c h1, .es-m-txt-c h2, .es-m-txt-c h3 { text-align:center!important } .es-m-txt-r, .es-m-txt-r h1, .es-m-txt-r h2, .es-m-txt-r h3 { text-align:right!important } .es-m-txt-l, .es-m-txt-l h1, .es-m-txt-l h2, .es-m-txt-l h3 { text-align:left!important } .es-m-txt-r img, .es-m-txt-c img, .es-m-txt-l img { display:inline!important } .es-button-border { display:inline-block!important } a.es-button, button.es-button { font-size:20px!important; display:inline-block!important } .es-adaptive table, .es-left, .es-right { width:100%!important } .es-content table, .es-header table, .es-footer table, .es-content, .es-footer, .es-header { width:100%!important; max-width:600px!important } .es-adapt-td { display:block!important; width:100%!important } .adapt-img { width:100%!important; height:auto!important } .es-m-p0 { padding:0!important } .es-m-p0r { padding-right:0!important } .es-m-p0l { padding-left:0!important } .es-m-p0t { padding-top:0!important } .es-m-p0b { padding-bottom:0!important } .es-m-p20b { padding-bottom:20px!important } .es-mobile-hidden, .es-hidden { display:none!important } tr.es-desk-hidden, td.es-desk-hidden, table.es-desk-hidden { width:auto!important; overflow:visible!important; float:none!important; max-height:inherit!important; line-height:inherit!important } tr.es-desk-hidden { display:table-row!important } table.es-desk-hidden { display:table!important } td.es-desk-menu-hidden { display:table-cell!important } .es-menu td { width:1%!important } table.es-table-not-adapt, .esd-block-html table { width:auto!important } table.es-social { display:inline-block!important } table.es-social td { display:inline-block!important } .es-m-p5 { padding:5px!important } .es-m-p5t { padding-top:5px!important } .es-m-p5b { padding-bottom:5px!important } .es-m-p5r { padding-right:5px!important } .es-m-p5l { padding-left:5px!important } .es-m-p10 { padding:10px!important } .es-m-p10t { padding-top:10px!important } .es-m-p10b { padding-bottom:10px!important } .es-m-p10r { padding-right:10px!important } .es-m-p10l { padding-left:10px!important } .es-m-p15 { padding:15px!important } .es-m-p15t { padding-top:15px!important } .es-m-p15b { padding-bottom:15px!important } .es-m-p15r { padding-right:15px!important } .es-m-p15l { padding-left:15px!important } .es-m-p20 { padding:20px!important } .es-m-p20t { padding-top:20px!important } .es-m-p20r { padding-right:20px!important } .es-m-p20l { padding-left:20px!important } .es-m-p25 { padding:25px!important } .es-m-p25t { padding-top:25px!important } .es-m-p25b { padding-bottom:25px!important } .es-m-p25r { padding-right:25px!important } .es-m-p25l { padding-left:25px!important } .es-m-p30 { padding:30px!important } .es-m-p30t { padding-top:30px!important } .es-m-p30b { padding-bottom:30px!important } .es-m-p30r { padding-right:30px!important } .es-m-p30l { padding-left:30px!important } .es-m-p35 { padding:35px!important } .es-m-p35t { padding-top:35px!important } .es-m-p35b { padding-bottom:35px!important } .es-m-p35r { padding-right:35px!important } .es-m-p35l { padding-left:35px!important } .es-m-p40 { padding:40px!important } .es-m-p40t { padding-top:40px!important } .es-m-p40b { padding-bottom:40px!important } .es-m-p40r { padding-right:40px!important } .es-m-p40l { padding-left:40px!important } .es-desk-hidden { display:table-row!important; width:auto!important; overflow:visible!important; max-height:inherit!important } }
@media screen and (max-width:384px) {.mail-message-content { width:414px!important } }
</style>
 </head>
 <body style="width:100%;font-family:arial, 'helvetica neue', helvetica, sans-serif;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding:0;Margin:0">
  <div dir="ltr" class="es-wrapper-color" lang="pt" style="background-color:#FAFAFA"><!--[if gte mso 9]>
			<v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t">
				<v:fill type="tile" color="#fafafa"></v:fill>
			</v:background>
		<![endif]-->
   <table class="es-wrapper" width="100%" cellspacing="0" cellpadding="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;padding:0;Margin:0;width:100%;height:100%;background-repeat:repeat;background-position:center top;background-color:#FAFAFA">
     <tr>
      <td valign="top" style="padding:0;Margin:0">
       <table cellpadding="0" cellspacing="0" class="es-header" align="center" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%;background-color:transparent;background-repeat:repeat;background-position:center top">
         <tr>
          <td align="center" style="padding:0;Margin:0">
           <table bgcolor="#ffffff" class="es-header-body" align="center" cellpadding="0" cellspacing="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:transparent;width:600px">
             <tr>
              <td align="left" style="Margin:0;padding-top:10px;padding-bottom:10px;padding-left:20px;padding-right:20px">
               <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                 <tr>
                  <td class="es-m-p0r" valign="top" align="center" style="padding:0;Margin:0;width:560px">
                   <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-bottom:20px;font-size:0px"><img src="https://ti.recordbrasilia.com.br/wp-content/uploads/2024/04/RECORD_BRASILIA_H_2D_PB_PS_RGB.png" alt="Logo" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;font-size:12px" width="200" title="Logo"></td>
                     </tr>
                   </table></td>
                 </tr>
               </table></td>
             </tr>
           </table></td>
         </tr>
       </table>
       <table cellpadding="0" cellspacing="0" class="es-content" align="center" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%">
         <tr>
          <td align="center" style="padding:0;Margin:0">
           <table bgcolor="#ffffff" class="es-content-body" align="center" cellpadding="0" cellspacing="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px">
             <tr>
              <td align="left" style="padding:0;Margin:0;padding-top:15px;padding-left:20px;padding-right:20px">
               <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                 <tr>
                  <td align="center" valign="top" style="padding:0;Margin:0;width:560px">
                   <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:10px;font-size:0px"><img src="https://ehcfpli.stripocdn.email/content/guids/CABINET_bc192ec9ea4283637813dcf6cc979b5e69f1a61b7a365202377869ae8e481204/images/botaorecsinaldegravacaovetoreps10isoladonofundobranco_3990893449_1.png" alt style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic" width="100"></td>
                     </tr>
                     <tr>
                      <td align="center" class="es-m-txt-c" style="padding:0;Margin:0;padding-top:15px;padding-bottom:15px"><h1 style="Margin:0;line-height:55px;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;font-size:46px;font-style:normal;font-weight:bold;color:#333333">Gravação Iniciada</h1></td>
                     </tr>
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:40px"><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;line-height:24px;color:#333333;font-size:16px">Olá, uma gravação na sala de reunião do 4º andar acaba de ser iniciada.</p></td>
                     </tr>
                   </table></td>
                 </tr>
               </table></td>
             </tr>
           </table></td>
         </tr>
       </table></td>
     </tr>
   </table>
  </div>
 </body>
</html>
"""
mensagem_inicio_programa = """\
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html dir="ltr" xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="pt">
 <head>
  <meta charset="UTF-8">
  <meta content="width=device-width, initial-scale=1" name="viewport">
  <meta name="x-apple-disable-message-reformatting">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta content="telephone=no" name="format-detection">
  <title>Novo modelo</title><!--[if (mso 16)]>
    <style type="text/css">
    a {text-decoration: none;}
    </style>
    <![endif]--><!--[if gte mso 9]><style>sup { font-size: 100% !important; }</style><![endif]--><!--[if gte mso 9]>
<xml>
    <o:OfficeDocumentSettings>
    <o:AllowPNG></o:AllowPNG>
    <o:PixelsPerInch>96</o:PixelsPerInch>
    </o:OfficeDocumentSettings>
</xml>
<![endif]-->
  <style type="text/css">
#outlook a {
	padding:0;
}
.es-button {
	mso-style-priority:100!important;
	text-decoration:none!important;
}
a[x-apple-data-detectors] {
	color:inherit!important;
	text-decoration:none!important;
	font-size:inherit!important;
	font-family:inherit!important;
	font-weight:inherit!important;
	line-height:inherit!important;
}
.es-desk-hidden {
	display:none;
	float:left;
	overflow:hidden;
	width:0;
	max-height:0;
	line-height:0;
	mso-hide:all;
}
@media only screen and (max-width:600px) {p, ul li, ol li, a { line-height:150%!important } h1, h2, h3, h1 a, h2 a, h3 a { line-height:120%!important } h1 { font-size:36px!important; text-align:left } h2 { font-size:26px!important; text-align:left } h3 { font-size:20px!important; text-align:left } .es-header-body h1 a, .es-content-body h1 a, .es-footer-body h1 a { font-size:36px!important; text-align:left } .es-header-body h2 a, .es-content-body h2 a, .es-footer-body h2 a { font-size:26px!important; text-align:left } .es-header-body h3 a, .es-content-body h3 a, .es-footer-body h3 a { font-size:20px!important; text-align:left } .es-menu td a { font-size:12px!important } .es-header-body p, .es-header-body ul li, .es-header-body ol li, .es-header-body a { font-size:14px!important } .es-content-body p, .es-content-body ul li, .es-content-body ol li, .es-content-body a { font-size:14px!important } .es-footer-body p, .es-footer-body ul li, .es-footer-body ol li, .es-footer-body a { font-size:14px!important } .es-infoblock p, .es-infoblock ul li, .es-infoblock ol li, .es-infoblock a { font-size:12px!important } *[class="gmail-fix"] { display:none!important } .es-m-txt-c, .es-m-txt-c h1, .es-m-txt-c h2, .es-m-txt-c h3 { text-align:center!important } .es-m-txt-r, .es-m-txt-r h1, .es-m-txt-r h2, .es-m-txt-r h3 { text-align:right!important } .es-m-txt-l, .es-m-txt-l h1, .es-m-txt-l h2, .es-m-txt-l h3 { text-align:left!important } .es-m-txt-r img, .es-m-txt-c img, .es-m-txt-l img { display:inline!important } .es-button-border { display:inline-block!important } a.es-button, button.es-button { font-size:20px!important; display:inline-block!important } .es-adaptive table, .es-left, .es-right { width:100%!important } .es-content table, .es-header table, .es-footer table, .es-content, .es-footer, .es-header { width:100%!important; max-width:600px!important } .es-adapt-td { display:block!important; width:100%!important } .adapt-img { width:100%!important; height:auto!important } .es-m-p0 { padding:0!important } .es-m-p0r { padding-right:0!important } .es-m-p0l { padding-left:0!important } .es-m-p0t { padding-top:0!important } .es-m-p0b { padding-bottom:0!important } .es-m-p20b { padding-bottom:20px!important } .es-mobile-hidden, .es-hidden { display:none!important } tr.es-desk-hidden, td.es-desk-hidden, table.es-desk-hidden { width:auto!important; overflow:visible!important; float:none!important; max-height:inherit!important; line-height:inherit!important } tr.es-desk-hidden { display:table-row!important } table.es-desk-hidden { display:table!important } td.es-desk-menu-hidden { display:table-cell!important } .es-menu td { width:1%!important } table.es-table-not-adapt, .esd-block-html table { width:auto!important } table.es-social { display:inline-block!important } table.es-social td { display:inline-block!important } .es-m-p5 { padding:5px!important } .es-m-p5t { padding-top:5px!important } .es-m-p5b { padding-bottom:5px!important } .es-m-p5r { padding-right:5px!important } .es-m-p5l { padding-left:5px!important } .es-m-p10 { padding:10px!important } .es-m-p10t { padding-top:10px!important } .es-m-p10b { padding-bottom:10px!important } .es-m-p10r { padding-right:10px!important } .es-m-p10l { padding-left:10px!important } .es-m-p15 { padding:15px!important } .es-m-p15t { padding-top:15px!important } .es-m-p15b { padding-bottom:15px!important } .es-m-p15r { padding-right:15px!important } .es-m-p15l { padding-left:15px!important } .es-m-p20 { padding:20px!important } .es-m-p20t { padding-top:20px!important } .es-m-p20r { padding-right:20px!important } .es-m-p20l { padding-left:20px!important } .es-m-p25 { padding:25px!important } .es-m-p25t { padding-top:25px!important } .es-m-p25b { padding-bottom:25px!important } .es-m-p25r { padding-right:25px!important } .es-m-p25l { padding-left:25px!important } .es-m-p30 { padding:30px!important } .es-m-p30t { padding-top:30px!important } .es-m-p30b { padding-bottom:30px!important } .es-m-p30r { padding-right:30px!important } .es-m-p30l { padding-left:30px!important } .es-m-p35 { padding:35px!important } .es-m-p35t { padding-top:35px!important } .es-m-p35b { padding-bottom:35px!important } .es-m-p35r { padding-right:35px!important } .es-m-p35l { padding-left:35px!important } .es-m-p40 { padding:40px!important } .es-m-p40t { padding-top:40px!important } .es-m-p40b { padding-bottom:40px!important } .es-m-p40r { padding-right:40px!important } .es-m-p40l { padding-left:40px!important } .es-desk-hidden { display:table-row!important; width:auto!important; overflow:visible!important; max-height:inherit!important } }
@media screen and (max-width:384px) {.mail-message-content { width:414px!important } }
</style>
 </head>
 <body style="width:100%;font-family:arial, 'helvetica neue', helvetica, sans-serif;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding:0;Margin:0">
  <div dir="ltr" class="es-wrapper-color" lang="pt" style="background-color:#FAFAFA"><!--[if gte mso 9]>
			<v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t">
				<v:fill type="tile" color="#fafafa"></v:fill>
			</v:background>
		<![endif]-->
   <table class="es-wrapper" width="100%" cellspacing="0" cellpadding="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;padding:0;Margin:0;width:100%;height:100%;background-repeat:repeat;background-position:center top;background-color:#FAFAFA">
     <tr>
      <td valign="top" style="padding:0;Margin:0">
       <table cellpadding="0" cellspacing="0" class="es-header" align="center" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%;background-color:transparent;background-repeat:repeat;background-position:center top">
         <tr>
          <td align="center" style="padding:0;Margin:0">
           <table bgcolor="#ffffff" class="es-header-body" align="center" cellpadding="0" cellspacing="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:transparent;width:600px">
             <tr>
              <td align="left" style="Margin:0;padding-top:10px;padding-bottom:10px;padding-left:20px;padding-right:20px">
               <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                 <tr>
                  <td class="es-m-p0r" valign="top" align="center" style="padding:0;Margin:0;width:560px">
                   <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-bottom:20px;font-size:0px"><img src="<caminho_img>" alt="Logo" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;font-size:12px" width="200" title="Logo"></td>
                     </tr>
                   </table></td>
                 </tr>
               </table></td>
             </tr>
           </table></td>
         </tr>
       </table>
       <table cellpadding="0" cellspacing="0" class="es-content" align="center" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%">
         <tr>
          <td align="center" style="padding:0;Margin:0">
           <table bgcolor="#ffffff" class="es-content-body" align="center" cellpadding="0" cellspacing="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px">
             <tr>
              <td align="left" style="padding:0;Margin:0;padding-top:15px;padding-left:20px;padding-right:20px">
               <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                 <tr>
                  <td align="center" valign="top" style="padding:0;Margin:0;width:560px">
                   <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:10px;font-size:0px"><img src="https://ehcfpli.stripocdn.email/content/guids/CABINET_bc192ec9ea4283637813dcf6cc979b5e69f1a61b7a365202377869ae8e481204/images/botaorecsinaldegravacaovetoreps10isoladonofundobranco_3990893449_1.png" alt style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic" width="100"></td>
                     </tr>
                     <tr>
                      <td align="center" class="es-m-txt-c" style="padding:0;Margin:0;padding-top:15px;padding-bottom:15px"><h1 style="Margin:0;line-height:55px;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;font-size:46px;font-style:normal;font-weight:bold;color:#333333">Programa Iniciado</h1></td>
                     </tr>
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:40px"><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;line-height:24px;color:#333333;font-size:16px">Olá, o programa foi inicializado.</p></td>
                     </tr>
                   </table></td>
                 </tr>
               </table></td>
             </tr>
           </table></td>
         </tr>
       </table></td>
     </tr>
   </table>
  </div>
 </body>
</html>
"""
mensagem_cancelamento = """\
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html dir="ltr" xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="pt">
 <head>
  <meta charset="UTF-8">
  <meta content="width=device-width, initial-scale=1" name="viewport">
  <meta name="x-apple-disable-message-reformatting">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta content="telephone=no" name="format-detection">
  <title>Novo modelo</title><!--[if (mso 16)]>
    <style type="text/css">
    a {text-decoration: none;}
    </style>
    <![endif]--><!--[if gte mso 9]><style>sup { font-size: 100% !important; }</style><![endif]--><!--[if gte mso 9]>
<xml>
    <o:OfficeDocumentSettings>
    <o:AllowPNG></o:AllowPNG>
    <o:PixelsPerInch>96</o:PixelsPerInch>
    </o:OfficeDocumentSettings>
</xml>
<![endif]-->
  <style type="text/css">
#outlook a {
	padding:0;
}
.es-button {
	mso-style-priority:100!important;
	text-decoration:none!important;
}
a[x-apple-data-detectors] {
	color:inherit!important;
	text-decoration:none!important;
	font-size:inherit!important;
	font-family:inherit!important;
	font-weight:inherit!important;
	line-height:inherit!important;
}
.es-desk-hidden {
	display:none;
	float:left;
	overflow:hidden;
	width:0;
	max-height:0;
	line-height:0;
	mso-hide:all;
}
@media only screen and (max-width:600px) {p, ul li, ol li, a { line-height:150%!important } h1, h2, h3, h1 a, h2 a, h3 a { line-height:120%!important } h1 { font-size:36px!important; text-align:left } h2 { font-size:26px!important; text-align:left } h3 { font-size:20px!important; text-align:left } .es-header-body h1 a, .es-content-body h1 a, .es-footer-body h1 a { font-size:36px!important; text-align:left } .es-header-body h2 a, .es-content-body h2 a, .es-footer-body h2 a { font-size:26px!important; text-align:left } .es-header-body h3 a, .es-content-body h3 a, .es-footer-body h3 a { font-size:20px!important; text-align:left } .es-menu td a { font-size:12px!important } .es-header-body p, .es-header-body ul li, .es-header-body ol li, .es-header-body a { font-size:14px!important } .es-content-body p, .es-content-body ul li, .es-content-body ol li, .es-content-body a { font-size:14px!important } .es-footer-body p, .es-footer-body ul li, .es-footer-body ol li, .es-footer-body a { font-size:14px!important } .es-infoblock p, .es-infoblock ul li, .es-infoblock ol li, .es-infoblock a { font-size:12px!important } *[class="gmail-fix"] { display:none!important } .es-m-txt-c, .es-m-txt-c h1, .es-m-txt-c h2, .es-m-txt-c h3 { text-align:center!important } .es-m-txt-r, .es-m-txt-r h1, .es-m-txt-r h2, .es-m-txt-r h3 { text-align:right!important } .es-m-txt-l, .es-m-txt-l h1, .es-m-txt-l h2, .es-m-txt-l h3 { text-align:left!important } .es-m-txt-r img, .es-m-txt-c img, .es-m-txt-l img { display:inline!important } .es-button-border { display:inline-block!important } a.es-button, button.es-button { font-size:20px!important; display:inline-block!important } .es-adaptive table, .es-left, .es-right { width:100%!important } .es-content table, .es-header table, .es-footer table, .es-content, .es-footer, .es-header { width:100%!important; max-width:600px!important } .es-adapt-td { display:block!important; width:100%!important } .adapt-img { width:100%!important; height:auto!important } .es-m-p0 { padding:0!important } .es-m-p0r { padding-right:0!important } .es-m-p0l { padding-left:0!important } .es-m-p0t { padding-top:0!important } .es-m-p0b { padding-bottom:0!important } .es-m-p20b { padding-bottom:20px!important } .es-mobile-hidden, .es-hidden { display:none!important } tr.es-desk-hidden, td.es-desk-hidden, table.es-desk-hidden { width:auto!important; overflow:visible!important; float:none!important; max-height:inherit!important; line-height:inherit!important } tr.es-desk-hidden { display:table-row!important } table.es-desk-hidden { display:table!important } td.es-desk-menu-hidden { display:table-cell!important } .es-menu td { width:1%!important } table.es-table-not-adapt, .esd-block-html table { width:auto!important } table.es-social { display:inline-block!important } table.es-social td { display:inline-block!important } .es-m-p5 { padding:5px!important } .es-m-p5t { padding-top:5px!important } .es-m-p5b { padding-bottom:5px!important } .es-m-p5r { padding-right:5px!important } .es-m-p5l { padding-left:5px!important } .es-m-p10 { padding:10px!important } .es-m-p10t { padding-top:10px!important } .es-m-p10b { padding-bottom:10px!important } .es-m-p10r { padding-right:10px!important } .es-m-p10l { padding-left:10px!important } .es-m-p15 { padding:15px!important } .es-m-p15t { padding-top:15px!important } .es-m-p15b { padding-bottom:15px!important } .es-m-p15r { padding-right:15px!important } .es-m-p15l { padding-left:15px!important } .es-m-p20 { padding:20px!important } .es-m-p20t { padding-top:20px!important } .es-m-p20r { padding-right:20px!important } .es-m-p20l { padding-left:20px!important } .es-m-p25 { padding:25px!important } .es-m-p25t { padding-top:25px!important } .es-m-p25b { padding-bottom:25px!important } .es-m-p25r { padding-right:25px!important } .es-m-p25l { padding-left:25px!important } .es-m-p30 { padding:30px!important } .es-m-p30t { padding-top:30px!important } .es-m-p30b { padding-bottom:30px!important } .es-m-p30r { padding-right:30px!important } .es-m-p30l { padding-left:30px!important } .es-m-p35 { padding:35px!important } .es-m-p35t { padding-top:35px!important } .es-m-p35b { padding-bottom:35px!important } .es-m-p35r { padding-right:35px!important } .es-m-p35l { padding-left:35px!important } .es-m-p40 { padding:40px!important } .es-m-p40t { padding-top:40px!important } .es-m-p40b { padding-bottom:40px!important } .es-m-p40r { padding-right:40px!important } .es-m-p40l { padding-left:40px!important } .es-desk-hidden { display:table-row!important; width:auto!important; overflow:visible!important; max-height:inherit!important } }
@media screen and (max-width:384px) {.mail-message-content { width:414px!important } }
</style>
 </head>
 <body style="width:100%;font-family:arial, 'helvetica neue', helvetica, sans-serif;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding:0;Margin:0">
  <div dir="ltr" class="es-wrapper-color" lang="pt" style="background-color:#FAFAFA"><!--[if gte mso 9]>
			<v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t">
				<v:fill type="tile" color="#fafafa"></v:fill>
			</v:background>
		<![endif]-->
   <table class="es-wrapper" width="100%" cellspacing="0" cellpadding="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;padding:0;Margin:0;width:100%;height:100%;background-repeat:repeat;background-position:center top;background-color:#FAFAFA">
     <tr>
      <td valign="top" style="padding:0;Margin:0">
       <table cellpadding="0" cellspacing="0" class="es-header" align="center" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%;background-color:transparent;background-repeat:repeat;background-position:center top">
         <tr>
          <td align="center" style="padding:0;Margin:0">
           <table bgcolor="#ffffff" class="es-header-body" align="center" cellpadding="0" cellspacing="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:transparent;width:600px">
             <tr>
              <td align="left" style="Margin:0;padding-top:10px;padding-bottom:10px;padding-left:20px;padding-right:20px">
               <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                 <tr>
                  <td class="es-m-p0r" valign="top" align="center" style="padding:0;Margin:0;width:560px">
                   <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-bottom:20px;font-size:0px"><img src="<caminho_img>" alt="Logo" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;font-size:12px" width="200" title="Logo"></td>
                     </tr>
                   </table></td>
                 </tr>
               </table></td>
             </tr>
           </table></td>
         </tr>
       </table>
       <table cellpadding="0" cellspacing="0" class="es-content" align="center" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%">
         <tr>
          <td align="center" style="padding:0;Margin:0">
           <table bgcolor="#ffffff" class="es-content-body" align="center" cellpadding="0" cellspacing="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px">
             <tr>
              <td align="left" style="padding:0;Margin:0;padding-top:15px;padding-left:20px;padding-right:20px">
               <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                 <tr>
                  <td align="center" valign="top" style="padding:0;Margin:0;width:560px">
                   <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:10px;font-size:0px"><img src="https://ehcfpli.stripocdn.email/content/guids/CABINET_bc192ec9ea4283637813dcf6cc979b5e69f1a61b7a365202377869ae8e481204/images/botaorecsinaldegravacaovetoreps10isoladonofundobranco_3990893449_1.png" alt style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic" width="100"></td>
                     </tr>
                     <tr>
                      <td align="center" class="es-m-txt-c" style="padding:0;Margin:0;padding-top:15px;padding-bottom:15px"><h1 style="Margin:0;line-height:55px;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;font-size:46px;font-style:normal;font-weight:bold;color:#333333">Gravação Cancelada</h1></td>
                     </tr>
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:40px"><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;line-height:24px;color:#333333;font-size:16px">Olá, entrevistador(a) optou por cancelar o zoom.</p></td>
                     </tr>
                   </table></td>
                 </tr>
               </table></td>
             </tr>
           </table></td>
         </tr>
       </table></td>
     </tr>
   </table>
  </div>
 </body>
</html>
"""
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
mouse_listener = pynput.mouse.Listener(suppress=True)
keyboard_listener = pynput.keyboard.Listener(suppress=True)


def copiar_arquivo(src, dst):
    SHGFP_TYPE_CURRENT = 0
    MAX_PATH = 260

    class FILEOPSTRUCT(ctypes.Structure):
        _fields_ = [("hwnd", wintypes.HWND),
                    ("wFunc", ctypes.wintypes.UINT),
                    ("pFrom", ctypes.c_wchar_p),
                    ("pTo", ctypes.c_wchar_p),
                    ("fFlags", ctypes.wintypes.INT),
                    ("fAnyOperationsAborted", ctypes.wintypes.BOOL),
                    ("hNameMappings", ctypes.c_void_p),
                    ("lpszProgressTitle", ctypes.c_wchar_p)]

    pFrom = src + '\0'
    pTo = dst + '\0'

    fileop = FILEOPSTRUCT()
    fileop.hwnd = None
    fileop.wFunc = 2
    fileop.pFrom = ctypes.c_wchar_p(pFrom)
    fileop.pTo = ctypes.c_wchar_p(pTo)
    fileop.fFlags = 0x0004
    fileop.fAnyOperationsAborted = False
    fileop.hNameMappings = None
    fileop.lpszProgressTitle = None

    ret = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(fileop))

    return ret == 0


def desabilitar_teclado_mouse():
    mouse_listener.start()
    keyboard_listener.start()


def habilitar_teclado_mouse():
    mouse_listener.stop()
    keyboard_listener.stop()


def inicio_programa(email_remetente, senha, email_destinatario_programa, assunto, mensagem_inicio_programa):

    smtp_server = "smtp-mail.outlook.com"
    smtp_port = 587

    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = email_destinatario_programa
    msg['Subject'] = assunto

    msg.attach(MIMEText(mensagem_inicio_programa, 'html'))

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_remetente, senha)

    server.sendmail(email_remetente, email_destinatario_programa, msg.as_string().encode('utf-8'))

    server.quit()


def inicio_gravacao(email_remetente, senha, email_destinatario, assunto, mensagem_inicio):

    smtp_server = "smtp-mail.outlook.com"
    smtp_port = 587

    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = " .".join(email_destinatario)
    msg['Subject'] = assunto

    msg.attach(MIMEText(mensagem_inicio, 'html'))

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_remetente, senha)

    server.sendmail(email_remetente, email_destinatario, msg.as_string().encode('utf-8'))

    server.quit()


def cancelamento_gravacao(email_remetente, senha, email_destinatario, assunto, mensagem_cancelamento):

    smtp_server = "smtp-mail.outlook.com"
    smtp_port = 587

    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = " .".join(email_destinatario)
    msg['Subject'] = assunto

    msg.attach(MIMEText(mensagem_cancelamento, 'html'))

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_remetente, senha)

    server.sendmail(email_remetente, email_destinatario, msg.as_string().encode('utf-8'))

    server.quit()


def conclusao_gravacao(email_remetente, senha, email_destinatario, assunto, mensagem_conclusao):

    smtp_server = "smtp-mail.outlook.com"
    smtp_port = 587

    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = " .".join(email_destinatario)
    msg['Subject'] = assunto

    msg.attach(MIMEText(mensagem_conclusao, 'html'))

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_remetente, senha)

    server.sendmail(email_remetente, email_destinatario, msg.as_string().encode('utf-8'))

    server.quit()


def enviar_contato(event=None):

    global entrada_contato
    entrada_contato = contato.get()
    if len(contato.get()) == 11 and contato.get().isnumeric():
        confirmacao = messagebox.askquestion('CONFIRMAÇÃO', f'O seguinte número: {entrada_contato} é o contato a ser enviado o convite?')
    else:
        messagebox.showerror('Número Inválido', 'Digite o número corretamente!')
        contato.delete(0, END)
        janela_contato.focus_set()
        contato_desselecionado(event)
           
    if confirmacao == 'no':
        contato.delete(0, END)
    else:
        janela_contato.destroy()


def contato_selecionado(event):
    if contato.get() == 'Ex: 61940028922':
        contato.delete(0, END)
        contato.configure(fg = "#000000")


def contato_desselecionado(event):
    if not contato.get():
        contato.insert(0, placeholder)
        contato.configure(fg = "#A9A9A9")


def abrir_pasta(caminho_pasta):
    subprocess.Popen(['explorer', caminho_pasta])


def salvar_retranca_r7():
    global nome
    global destino
    global pasta
    origem = f'C:/Users/{usuario_atual}/videos/ZOOM.mkv'
    nome = retranca.get().upper()
    destino = '<caminho>'.format(nome)
    pasta = '<caminho>'
    
    if len(nome) == 0:
        messagebox.showerror('ERRO: NOME DO ARQUIVO', 'Digite a retranca corretamente!')

    else:
        confirma_nome_destino_gravacao = messagebox.askquestion("CONFIRMAÇÃO RETRANCA", f'Você quer que o nome do arquivo seja: "{nome}"?')

    if confirma_nome_destino_gravacao == 'yes':
        janela_retranca.destroy()
        speak('Aguarde enquanto envio a gravação para a pasta.')
        diretorio, nome_arquivo = os.path.split(origem)
        nome_extensao = nome + '.mkv'
        novo_nome = os.path.join(diretorio, nome_extensao)
        success = copiar_arquivo(origem, destino)
        os.rename(origem, novo_nome)
        speak('Certo, conclui o envio. Vou abrir a pasta para mostrar.')
        caminho_pasta = '<caminho>'
        abrir_pasta(caminho_pasta)
        desabilitar_teclado_mouse()
        time.sleep(10)
        habilitar_teclado_mouse()
        pag.hotkey('alt', 'F4')

    else:
        retranca.delete(0, END)


def abrir_pasta_local(caminho_pasta_local):
    subprocess.Popen(['explorer', caminho_pasta_local])


def salvar_retranca_local():
    global nome
    global destino
    global pasta
    origem = '<caminho>'
    nome = retranca.get().upper()
    destino = '<caminho>'.format(nome)
    pasta = '<caminho>'
    

    if len(nome) == 0:
        messagebox.showerror('ERRO: NOME DO ARQUIVO', 'Digite a retranca corretamente!')

    else:
        confirma_nome_destino_gravacao = messagebox.askquestion("CONFIRMAÇÃO RETRANCA", f'Você quer que o nome do arquivo seja: "{nome}"?')

    if confirma_nome_destino_gravacao == 'yes':
        janela_retranca.destroy()
        speak('Aguarde enquanto envio a gravação para a pasta.')
        diretorio, nome_arquivo = os.path.split(origem)
        nome_extensao = nome + '.mkv'
        novo_nome = os.path.join(diretorio, nome_extensao)
        success = copiar_arquivo(origem, destino)
        os.rename(origem, novo_nome)
        speak('Certo, conclui o envio. Vou abrir a pasta para mostrar.')
        caminho_pasta_local = '<caminho>'
        abrir_pasta_local(caminho_pasta_local)
        desabilitar_teclado_mouse()
        time.sleep(10)
        habilitar_teclado_mouse()
        pag.hotkey('alt', 'F4')

    else:
        retranca.delete(0, END)


def abrir_pasta_nacional(caminho_pasta_nacional):
    subprocess.Popen(['explorer', caminho_pasta_nacional])


def salvar_retranca_nacional():
    global nome
    global destino
    global pasta
    origem = '<caminho>'
    nome = retranca.get().upper()
    destino = '<caminho>'.format(nome)
    pasta = '<caminho>'
    

    if len(nome) == 0:
        messagebox.showerror('ERRO: NOME DO ARQUIVO', 'Digite a retranca corretamente!')

    else:
        confirma_nome_destino_gravacao = messagebox.askquestion("CONFIRMAÇÃO RETRANCA", f'Você quer que o nome do arquivo seja: "{nome}"?')

    if confirma_nome_destino_gravacao == 'yes':
        janela_retranca.destroy()
        speak('Aguarde enquanto envio a gravação para a pasta do jornalismo nacional.')
        diretorio, nome_arquivo = os.path.split(origem)
        nome_extensao = nome + '.mkv'
        novo_nome = os.path.join(diretorio, nome_extensao)
        success = copiar_arquivo(origem, destino)
        os.rename(origem, nome)
        speak('Certo, conclui o envio. Vou abrir a pasta para mostrar.')
        caminho_pasta_nacional = '<caminho>'
        abrir_pasta_nacional(caminho_pasta_nacional)
        desabilitar_teclado_mouse()
        time.sleep(10)
        habilitar_teclado_mouse()
        pag.hotkey('alt', 'F4')

    else:
        retranca.delete(0, END)


def speak(text):
    engine.say(text)
    engine.runAndWait()


def update_animation_volume():
    global current_frame_volume
    current_frame_volume += 1
    if current_frame_volume >= len(frames_volume):
        current_frame_volume = 0
    canvas.itemconfig(background, image=frames_volume[current_frame_volume])
    janela_volume.after(animation_speed_volume, update_animation_volume)


def load_gif_frames_volume():
    global frames_volume
    fundo_volume = imageio.mimread('<caminho>')
    frames_volume = [PhotoImage(format="gif", data=frame.data, master=janela_volume) for frame in fundo_volume]


def fechar_janela_volume():
    janela_volume.destroy()


def btn_clicked():
    speak('Certo, irei encerrar o programa.')
    window.destroy()
    pag.hotkey('alt', 'q')
    pag.press('enter')
    pag.hotkey('alt', 'F4')
    cancelamento_gravacao(email_remetente, senha, email_destinatario_gravacao, assunto, mensagem_cancelamento)
    # while True:
    #     try:
    #         response = requests.get('{}/win&A=0'.format(URL_LED))
    #         break
    #     except Exception:
    #         break
    logout_windows()
    sys.exit()


def cancelamento_gravacao_contato():
    speak('Certo, irei encerrar o programa.')
    janela_contato.destroy()
    pag.hotkey('alt', 'q')
    pag.press('enter')
    pag.hotkey('alt', 'F4')
    cancelamento_gravacao(email_remetente, senha, email_destinatario_gravacao, assunto, mensagem_cancelamento)
    # while True:
    #     try:
    #         response = requests.get('{}/win&A=0'.format(URL_LED))
    #         break
    #     except Exception:
    #         break
    logout_windows()
    sys.exit()


def logout_windows():
    EWX_LOGOFF = 0x00000000
    EWX_FORCE = 0x00000004

    ctypes.windll.user32.ExitWindowsEx(EWX_LOGOFF | EWX_FORCE, 0)


pag.PAUSE = 2

inicio_programa(email_remetente, senha, email_destinatario_programa, assunto, mensagem_inicio_programa)

# while True:
#     try:
#         response = requests.get('{}/win&A=255&R=255&G=170&B=0&FX=0'.format(URL_LED), timeout=3)
#         break
#     except Exception:
#         break

janela_volume = Tk()
largura_janela = 700
altura_janela = 700

largura_tela = janela_volume.winfo_screenwidth()
altura_tela = janela_volume.winfo_screenheight()

pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

janela_volume.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
janela_volume.title("Volume TV")

icone_volume = Image.open('<caminho>')
photo_volume = ImageTk.PhotoImage(icone_volume)
janela_volume.wm_iconphoto(False, photo_volume)

janela_volume.configure(bg="#ffffff")

canvas = Canvas(
    janela_volume,
    bg="#ffffff",
    height=altura_janela,
    width=largura_janela,
    bd=0,
    highlightthickness=0,
    relief="ridge")
canvas.place(x=0, y=0)

fundo_volume = Image.open('<caminho>')
frames_volume = [ImageTk.PhotoImage(frame) for frame in ImageSequence.Iterator(fundo_volume)]

current_frame_volume = 0
background = canvas.create_image(
    largura_janela / 2, altura_janela / 2,
    anchor="center")

animation_speed_volume = 40
frames_volume_per_check = 40

thread = threading.Thread(target=load_gif_frames_volume)
thread.daemon = True
thread.start()

current_frame_volume = 0
background = canvas.create_image(
    largura_janela / 2, altura_janela / 2,
    anchor="center")

janela_volume.resizable(False, False)

update_animation_volume()

janela_volume.after(5000, fechar_janela_volume)

janela_volume.mainloop()

speak('Vou abrir os programas necessários.')

pag.press('win')
keyboard.write('zoom')
time.sleep(0.5)
pag.press('enter')

while 1:
    try:
        pag.locateOnScreen('<caminho>', confidence=0.8)
        x, y, largura, altura = pag.locateOnScreen('<caminho>')
        pag.click(x + largura/2, y + altura/2)
        break
    except pag.ImageNotFoundException:
        continue
    time.sleep(1)


while 1:
    try:
        pag.locateOnScreen('<caminho>', confidence=0.8)
        pag.hotkey('alt', 'shift', 'i')
        break
    except pag.ImageNotFoundException:
        continue
    
pag.hotkey('win', 'up')

janela_contato = Tk()
largura_janela = 700
altura_janela = 700

largura_tela = janela_contato.winfo_screenwidth()
altura_tela = janela_contato.winfo_screenheight()

pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

janela_contato.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
janela_contato.title('Contato')

icone_contato = Image.open('<caminho>')
photo_contato = ImageTk.PhotoImage(icone_contato)
janela_contato.wm_iconphoto(False, photo_contato)

janela_contato.configure(bg = "#ffffff")
canvas = Canvas(
    janela_contato,
    bg = "#ffffff",
    height = altura_janela,
    width = largura_janela,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

img_fundo_contato = PhotoImage(file = '<caminho>')
fundo_contato = canvas.create_image(
    largura_janela / 2, altura_janela / 2,
    image=img_fundo_contato)

entry0_img = PhotoImage(file = '<caminho>')
entry0_bg = canvas.create_image(
    335.0, 358.0,
    image = entry0_img)

contato = Entry(
    bd = 0,
    bg = "#d9d9d9",
    font = ('Record Type', 14),
    highlightthickness = 0)

contato.place(
    x = 111, y = 340,
    width = 448,
    height = 34)

img0 = PhotoImage(file = '<caminho>')
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = enviar_contato,
    relief = "flat",
    cursor = 'hand2')

b0.place(
    x = 180, y = 455,
    width = 136,
    height = 40)

img1 = PhotoImage(file = '<caminho>')
b1 = Button(
    image = img1,
    borderwidth = 0,
    highlightthickness = 0,
    command = cancelamento_gravacao_contato,
    relief = "flat",
    cursor= 'hand2')

b1.place(
    x = 363, y = 456,
    width = 157,
    height = 37)

placeholder = 'Ex: 61940028922'
contato.insert(0, placeholder)
contato.bind("<FocusIn>", contato_selecionado)
contato.bind('<FocusOut>', contato_desselecionado)
b0.bind('<ButtonRelease-1>', lambda event: enviar_contato(event))
janela_contato.bind('<Return>', enviar_contato)
janela_contato.resizable(False, False)
janela_contato.attributes('-topmost', True)
janela_contato.mainloop()

convite = pyperclip.paste()
mensagem_wpp = {
    "number": f"[55{entrada_contato}]",
    "textMessage": {
        "text": f"Olá, você está recebendo o *link para a entrevista com a Record Brasília:*\n\n{convite}\n\n*Caso o link não esteja disponível adicione este contato para o link ser ativado ou, responda com um 'Oi' saia da conversa e entre na conversa novamente.*"
    }
}

api_key = "<chave>"
headers = {
    "apikey": f"{api_key}",
    "Content-Type": "application/json"
}

speak('Irei enviar o link do zoom para o contato.')

response = requests.post(URL_WPP, json=mensagem_wpp, headers=headers)

speak('Link do zoom foi enviado para o contato.')

pag.hotkey('alt', 'u')

def update_animation():
    global current_frame
    current_frame += 1
    if current_frame >= len(frames):
        current_frame = 0
    canvas.itemconfig(background, image=frames[current_frame])
    window.after(animation_speed, update_animation)

    if current_frame % frames_per_check == 0:
        check_image_on_screen()

def check_image_on_screen():
    global current_image
    try:
        pag.locateOnScreen('<caminho>', confidence=0.8)
        window.destroy()
    except pag.ImageNotFoundException:
        pass

def load_gif_frames():
    global frames
    fundo_aguardando = imageio.mimread('<caminho>')
    frames = [PhotoImage(format="gif", data=frame.data, master=window) for frame in fundo_aguardando]

window = Tk()
largura_janela = 700
altura_janela = 700

largura_tela = window.winfo_screenwidth()
altura_tela = window.winfo_screenheight()

pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

window.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
window.title('Aguardando')

icone_animacao = Image.open('<caminho>')
photo_animacao = ImageTk.PhotoImage(icone_animacao)
window.wm_iconphoto(False, photo_animacao)

window.configure(bg="#ffffff")

canvas = Canvas(
    window,
    bg="#ffffff",
    height=altura_janela,
    width=largura_janela,
    bd=0,
    highlightthickness=0,
    relief="ridge")
canvas.place(x=0, y=0)

fundo_aguardando = Image.open('<caminho>')
frames = [ImageTk.PhotoImage(frame) for frame in ImageSequence.Iterator(fundo_aguardando)]

current_frame = 0
background = canvas.create_image(
    largura_janela / 2, altura_janela / 2,
    anchor="center")

animation_speed = 40
frames_per_check = 45

thread = threading.Thread(target=load_gif_frames)
thread.daemon = True
thread.start()

current_frame = 0
background = canvas.create_image(
    largura_janela / 2, altura_janela / 2,
    anchor="center")

img_cancelar_reuniao = PhotoImage(file='<caminho>')
botao_cancelar_reuniao = Button(
    image=img_cancelar_reuniao,
    borderwidth=0,
    highlightthickness=0,
    command=btn_clicked,
    relief="flat",
    cursor='hand2')
botao_cancelar_reuniao.place(
    x=258, y=569,
    width=184,
    height=29)

window.resizable(False, False)

update_animation()

window.mainloop()

while 1:
    try:
        pag.locateOnScreen('<caminho>', confidence=0.8)
        pag.press('win')
        keyboard.write('obs')
        time.sleep(0.5)
        pag.press('enter') 
        break
    except pag.ImageNotFoundException:
        continue
    time.sleep(1)

while 1:
    try:
        pag.locateOnScreen('<caminho>', confidence=0.8)
        pag.press('shift')
        speak('Gravação iniciada.')
        pag.hotkey('alt', 'space')
        pag.press('n')
        break
    except pag.ImageNotFoundException:
        continue
    time.sleep(1)

# while True:
#     try:
#         response = requests.get('{}/win&A=255&R=255&G=&B=&FX=12'.format(URL_LED), timeout=3)
#         break
#     except Exception:
#         break

inicio_gravacao(email_remetente, senha, email_destinatario_gravacao, assunto, mensagem_inicio)

pag.hotkey('alt', 'u')
pag.moveTo(1919, 367)

keyboard.wait('space')
speak('Encerrando gravação.')

pag.hotkey('alt', 'q')
pag.press('enter')

time.sleep(0.5)

while 1:
    try:
        pag.hotkey('alt', 'shift', 'tab')
        pag.locateOnScreen('<caminho>', confidence=0.8)
        pag.hotkey('win', 'up')
        pag.press('shift')
        time.sleep(0.5)
        pag.hotkey('alt', 'F4')
        break
    except pag.ImageNotFoundException:
        continue
    time.sleep(1)

pag.hotkey('alt', 'F4')

# while True:
#     try:
#         response = requests.get('{}/win&A=255&R=&G=255&B=&FX=0'.format(URL_LED), timeout=3)
#         break
#     except Exception:
#         break
    
speak('Gravação concluída.')

janela_retranca = Tk()
largura_janela = 700
altura_janela = 700

largura_tela = janela_retranca.winfo_screenwidth()
altura_tela = janela_retranca.winfo_screenheight()

pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

janela_retranca.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
janela_retranca.title('Nome Arquivo de Gravação')

icone_retranca = Image.open('<caminho>')
photo_retranca = ImageTk.PhotoImage(icone_retranca)
janela_retranca.wm_iconphoto(True, photo_retranca)

janela_retranca.configure(bg = "#ffffff")
canvas = Canvas(
    janela_retranca,
    bg = "#ffffff",
    height = altura_janela,
    width = largura_janela,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

img_fundo_janela_retranca = PhotoImage(file = '<caminho>')
fundo_janela_retranca = canvas.create_image(
    largura_janela / 2, altura_janela / 2,
    image=img_fundo_janela_retranca)

img_entrada_retranca = PhotoImage(file = '<caminho>')
entrada_retranca = canvas.create_image(
    349.5, 331.0,
    image = img_entrada_retranca)

retranca = Entry(
    bd = 0,
    bg = "#d9d9d9",
    font = ('Record Type', 14),
    highlightthickness = 0)

retranca.place(
    x = 117, y = 312,
    width = 465,
    height = 36)

img_botao_local = PhotoImage(file = '<caminho>')
botao_local = Button(
    image = img_botao_local,
    borderwidth = 0,
    highlightthickness = 0,
    command = salvar_retranca_local,
    relief = "flat",
    cursor = 'hand2')

botao_local.place(
    x = 281, y = 469,
    width = 89,
    height = 41)

img_botao_r7 = PhotoImage(file = '<caminho>')
botao_r7 = Button(
    image = img_botao_r7,
    borderwidth = 0,
    highlightthickness = 0,
    command = salvar_retranca_r7,
    relief = "flat",
    cursor = 'hand2')

botao_r7.place(
    x = 149, y = 469,
    width = 90,
    height = 42)

img_botao_nacional = PhotoImage(file = '<caminho>')
botao_nacional = Button(
    image = img_botao_nacional,
    borderwidth = 0,
    highlightthickness = 0,
    command = salvar_retranca_nacional,
    relief = "flat",
    cursor = 'hand2')

botao_nacional.place(
    x = 413, y = 470,
    width = 138,
    height = 40)

janela_retranca.resizable(False, False)
janela_retranca.attributes('-topmost', True)
janela_retranca.mainloop()

mensagem_conclusao = """\
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html dir="ltr" xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="pt">
 <head>
  <meta charset="UTF-8">
  <meta content="width=device-width, initial-scale=1" name="viewport">
  <meta name="x-apple-disable-message-reformatting">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta content="telephone=no" name="format-detection">
  <title>Novo modelo</title><!--[if (mso 16)]>
    <style type="text/css">
    a {text-decoration: none;}
    </style>
    <![endif]--><!--[if gte mso 9]><style>sup { font-size: 100% !important; }</style><![endif]--><!--[if gte mso 9]>
<xml>
    <o:OfficeDocumentSettings>
    <o:AllowPNG></o:AllowPNG>
    <o:PixelsPerInch>96</o:PixelsPerInch>
    </o:OfficeDocumentSettings>
</xml>
<![endif]-->
  <style type="text/css">
#outlook a {
	padding:0;
}
.es-button {
	mso-style-priority:100!important;
	text-decoration:none!important;
}
a[x-apple-data-detectors] {
	color:inherit!important;
	text-decoration:none!important;
	font-size:inherit!important;
	font-family:inherit!important;
	font-weight:inherit!important;
	line-height:inherit!important;
}
.es-desk-hidden {
	display:none;
	float:left;
	overflow:hidden;
	width:0;
	max-height:0;
	line-height:0;
	mso-hide:all;
}
@media only screen and (max-width:600px) {p, ul li, ol li, a { line-height:150%!important } h1, h2, h3, h1 a, h2 a, h3 a { line-height:120%!important } h1 { font-size:36px!important; text-align:left } h2 { font-size:26px!important; text-align:left } h3 { font-size:20px!important; text-align:left } .es-header-body h1 a, .es-content-body h1 a, .es-footer-body h1 a { font-size:36px!important; text-align:left } .es-header-body h2 a, .es-content-body h2 a, .es-footer-body h2 a { font-size:26px!important; text-align:left } .es-header-body h3 a, .es-content-body h3 a, .es-footer-body h3 a { font-size:20px!important; text-align:left } .es-menu td a { font-size:12px!important } .es-header-body p, .es-header-body ul li, .es-header-body ol li, .es-header-body a { font-size:14px!important } .es-content-body p, .es-content-body ul li, .es-content-body ol li, .es-content-body a { font-size:14px!important } .es-footer-body p, .es-footer-body ul li, .es-footer-body ol li, .es-footer-body a { font-size:14px!important } .es-infoblock p, .es-infoblock ul li, .es-infoblock ol li, .es-infoblock a { font-size:12px!important } *[class="gmail-fix"] { display:none!important } .es-m-txt-c, .es-m-txt-c h1, .es-m-txt-c h2, .es-m-txt-c h3 { text-align:center!important } .es-m-txt-r, .es-m-txt-r h1, .es-m-txt-r h2, .es-m-txt-r h3 { text-align:right!important } .es-m-txt-l, .es-m-txt-l h1, .es-m-txt-l h2, .es-m-txt-l h3 { text-align:left!important } .es-m-txt-r img, .es-m-txt-c img, .es-m-txt-l img { display:inline!important } .es-button-border { display:inline-block!important } a.es-button, button.es-button { font-size:20px!important; display:inline-block!important } .es-adaptive table, .es-left, .es-right { width:100%!important } .es-content table, .es-header table, .es-footer table, .es-content, .es-footer, .es-header { width:100%!important; max-width:600px!important } .es-adapt-td { display:block!important; width:100%!important } .adapt-img { width:100%!important; height:auto!important } .es-m-p0 { padding:0!important } .es-m-p0r { padding-right:0!important } .es-m-p0l { padding-left:0!important } .es-m-p0t { padding-top:0!important } .es-m-p0b { padding-bottom:0!important } .es-m-p20b { padding-bottom:20px!important } .es-mobile-hidden, .es-hidden { display:none!important } tr.es-desk-hidden, td.es-desk-hidden, table.es-desk-hidden { width:auto!important; overflow:visible!important; float:none!important; max-height:inherit!important; line-height:inherit!important } tr.es-desk-hidden { display:table-row!important } table.es-desk-hidden { display:table!important } td.es-desk-menu-hidden { display:table-cell!important } .es-menu td { width:1%!important } table.es-table-not-adapt, .esd-block-html table { width:auto!important } table.es-social { display:inline-block!important } table.es-social td { display:inline-block!important } .es-m-p5 { padding:5px!important } .es-m-p5t { padding-top:5px!important } .es-m-p5b { padding-bottom:5px!important } .es-m-p5r { padding-right:5px!important } .es-m-p5l { padding-left:5px!important } .es-m-p10 { padding:10px!important } .es-m-p10t { padding-top:10px!important } .es-m-p10b { padding-bottom:10px!important } .es-m-p10r { padding-right:10px!important } .es-m-p10l { padding-left:10px!important } .es-m-p15 { padding:15px!important } .es-m-p15t { padding-top:15px!important } .es-m-p15b { padding-bottom:15px!important } .es-m-p15r { padding-right:15px!important } .es-m-p15l { padding-left:15px!important } .es-m-p20 { padding:20px!important } .es-m-p20t { padding-top:20px!important } .es-m-p20r { padding-right:20px!important } .es-m-p20l { padding-left:20px!important } .es-m-p25 { padding:25px!important } .es-m-p25t { padding-top:25px!important } .es-m-p25b { padding-bottom:25px!important } .es-m-p25r { padding-right:25px!important } .es-m-p25l { padding-left:25px!important } .es-m-p30 { padding:30px!important } .es-m-p30t { padding-top:30px!important } .es-m-p30b { padding-bottom:30px!important } .es-m-p30r { padding-right:30px!important } .es-m-p30l { padding-left:30px!important } .es-m-p35 { padding:35px!important } .es-m-p35t { padding-top:35px!important } .es-m-p35b { padding-bottom:35px!important } .es-m-p35r { padding-right:35px!important } .es-m-p35l { padding-left:35px!important } .es-m-p40 { padding:40px!important } .es-m-p40t { padding-top:40px!important } .es-m-p40b { padding-bottom:40px!important } .es-m-p40r { padding-right:40px!important } .es-m-p40l { padding-left:40px!important } .es-desk-hidden { display:table-row!important; width:auto!important; overflow:visible!important; max-height:inherit!important } .h-auto { height:auto!important } }
@media screen and (max-width:384px) {.mail-message-content { width:414px!important } }
</style>
 </head>
 <body style="width:100%;font-family:arial, 'helvetica neue', helvetica, sans-serif;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding:0;Margin:0">
  <div dir="ltr" class="es-wrapper-color" lang="pt" style="background-color:#FAFAFA"><!--[if gte mso 9]>
			<v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t">
				<v:fill type="tile" color="#fafafa"></v:fill>
			</v:background>
		<![endif]-->
   <table class="es-wrapper" width="100%" cellspacing="0" cellpadding="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;padding:0;Margin:0;width:100%;height:100%;background-repeat:repeat;background-position:center top;background-color:#FAFAFA">
     <tr>
      <td valign="top" style="padding:0;Margin:0">
       <table cellpadding="0" cellspacing="0" class="es-header" align="center" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%;background-color:transparent;background-repeat:repeat;background-position:center top">
         <tr>
          <td align="center" style="padding:0;Margin:0">
           <table bgcolor="#ffffff" class="es-header-body" align="center" cellpadding="0" cellspacing="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:transparent;width:600px">
             <tr>
              <td align="left" style="Margin:0;padding-top:10px;padding-bottom:10px;padding-left:20px;padding-right:20px">
               <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                 <tr>
                  <td class="es-m-p0r" valign="top" align="center" style="padding:0;Margin:0;width:560px">
                   <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-bottom:20px;font-size:0px"><img src="'<caminho>'" alt="Logo" style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;font-size:12px" width="200" title="Logo"></td>
                     </tr>
                   </table></td>
                 </tr>
               </table></td>
             </tr>
           </table></td>
         </tr>
       </table>
       <table cellpadding="0" cellspacing="0" class="es-content" align="center" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;table-layout:fixed !important;width:100%">
         <tr>
          <td align="center" style="padding:0;Margin:0">
           <table bgcolor="#ffffff" class="es-content-body" align="center" cellpadding="0" cellspacing="0" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px">
             <tr>
              <td align="left" style="padding:0;Margin:0;padding-top:15px;padding-left:20px;padding-right:20px">
               <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                 <tr>
                  <td align="center" valign="top" style="padding:0;Margin:0;width:560px">
                   <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:10px;font-size:0px"><img src="https://ehcfpli.stripocdn.email/content/guids/CABINET_bc192ec9ea4283637813dcf6cc979b5e69f1a61b7a365202377869ae8e481204/images/botaorecsinaldegravacaovetoreps10isoladonofundobranco_3990893449_1.png" alt style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic" width="100"></td>
                     </tr>
                     <tr>
                      <td align="center" class="es-m-txt-c" style="padding:0;Margin:0;padding-top:15px;padding-bottom:15px"><h1 style="Margin:0;line-height:55px;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;font-size:46px;font-style:normal;font-weight:bold;color:#333333">Gravação Encerrada</h1></td>
                     </tr>
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px;padding-bottom:10px"><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;line-height:24px;color:#333333;font-size:16px">Olá, a gravação na sala de reunião do 4º andar acaba de ser encerrada.</p></td>
                     </tr>
                     <tr>
                      <td align="center" style="padding:0;Margin:0;padding-top:10px"><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;line-height:24px;color:#333333;font-size:16px">"""f"""\
                      A retranca <strong>{nome}</strong>, está salva no caminho de rede abaixo:</p><p><strong>{pasta}</strong>&nbsp;</p><p style="Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:arial, 'helvetica neue', helvetica, sans-serif;line-height:24px;color:#333333;font-size:16px"><br></p></td>
                     </tr>
                   </table></td>
                 </tr>
               </table></td>
             </tr>
           </table></td>
         </tr>
       </table></td>
     </tr>
   </table>
  </div>
 </body>
</html>
"""

conclusao_gravacao(email_remetente, senha, email_destinatario_gravacao, assunto, mensagem_conclusao)

# while True:
#     try:
#       response = requests.get('{}/win&A=0'.format(URL_LED), timeout=3)
#       break
#     except Exception:
#         break
    
speak('Programa será fechado e irei fazer logoff.')
logout_windows()