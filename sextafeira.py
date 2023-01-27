from http.client import ImproperConnectionState
from importlib.resources import path
from logging import exception
import pyttsx3 #bibliotece de fala
import datetime #bibliotece de data  e hora
import speech_recognition as sr #bibliotece de  reconhecimento voz
import wikipedia
import time
import os.path
import webbrowser
import subprocess
import ctypes
import win32com.client as wincl 


texto_fala = pyttsx3.init()
chrome_path = 'C:\Program Files\Google\Chrome\Application\chrome %s'

def falar(audio):
    rate = texto_fala.getProperty('rate') 
    texto_fala.setProperty('rate',240) #alterando a velocidade da voz

    voices = texto_fala.getProperty('voices')
    texto_fala.setProperty('voice',voices[0].id) #alteração do tipo de voz da assistente
    texto_fala.say(audio)
    texto_fala.runAndWait()

   

#reconhecendo a voz
def ouvir_microfone():
    microfone = sr.Recognizer()

    try:
        
        with sr.Microphone() as source:
            microfone.adjust_for_ambient_noise(source)
            audio = microfone.listen(source)
            print('Reconhecendo...')
            comando = microfone.recognize_google(audio,language='pt-br')
            time.sleep(1)
            print(comando)
            texto_fala.runAndWait()

    except Exception as e:
        print(e)
        falar('Não entendi, por favor repita!')

        return "none"

    return comando

if __name__ == "__main__":
    limpar = lambda: os.system('cls')
    
    limpar()
       
    while True:
        comando = ouvir_microfone().lower()
        texto_fala.runAndWait() 

        if 'horas' in comando:
            hora = datetime.datetime.now().strftime("%H:%M")
            texto_fala.say('Agora são:'+ hora)
            texto_fala.runAndWait()
        
        elif 'abrir youtube' in comando:
            texto_fala.say("Abrindo Youtube\n")
            webbrowser.get("C:\Program Files\Google\Chrome\Application\chrome %s")
            webbrowser.open("youtube.com.br")
            texto_fala.runAndWait()
 
        elif 'abrir navegador' in comando or "abrir google" in comando:
            texto_fala.say("Abrindo o navegador Google")
            webbrowser.open("google.com.br")
            texto_fala.runAndWait()
        


        if ' abrir whatsApp' in comando:
            texto_fala.say("certo. abrindo whatsapp")
            path = "C:\Users\TIAGO\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\WhatsApp"
            subprocess.run(path)
            texto_fala.runAndWait()

        
        elif 'como você está' in comando:
            texto_fala.say("Eu estou bem, obrigado! e como você está?")
            texto_fala.runAndWait()
 
        elif 'bem' in comando or "bom" in comando:
            texto_fala.say("que bom ,que esta tudo bem com você")
            texto_fala.runAndWait()
        
        elif 'encerrar' in comando:
            time.sleep(1)
            texto_fala.say("tudo bem")
            exit()
            
        elif 'bloquear computador' in comando:
            texto_fala.say("computador bloqueado com sucesso!")
            ctypes.windll.user32.LockWorkStation()
            time.sleep(1)
            texto_fala.runAndWait()
            
        elif "log off" in comando or "trocar de usuário" in comando:
            texto_fala.say(" tudo bem! Apenas certifique-se de que todos os aplicativos estejam fechados antes de sair")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"]) 

 
        elif 'desligar sistema' in comando:
                texto_fala.say("Espere um segundo ! Seu sistema irá desligar")
                subprocess.call('shutdown / p /f')

        elif 'sexta-feira' in comando:
            hora = datetime.datetime.now().hour
            if hora >= 6 and hora < 12:
                falar('Bom dia, Como posso ajudar?')
            elif hora >=12 and hora <18:
                falar('Boa tarde, Como posso ajudar?')
            elif hora >=18 and hora <24:
                falar('Boa noite. Como posso ajudar?')
            else:
                ('')
            texto_fala.runAndWait()


        elif 'dia' in comando:
            time.sleep(1)
            dia =int( datetime.datetime.now().day)
            mes =int( datetime.datetime.now().month)
            ano =int( datetime.datetime.now().year)
            texto_fala.say('Hoje são')
            texto_fala.say(dia)
            texto_fala.say('do')
            texto_fala.say(mes)
            texto_fala.say('de')
            texto_fala.say(ano)
            texto_fala.runAndWait()

        elif 'pesquise por' in comando:
            time.sleep(1)
            procurar = comando.replace('pesquise por','')
            wikipedia.set_lang('pt')
            resultado = wikipedia.summary(procurar,2)
            texto_fala.say(resultado)
            texto_fala.runAndWait()
            
            
        elif 'reproduzir' in comando:
            musica = comando.replace('reproduzir','')
            texto_fala.say('tocando'+ musica)
            texto_fala.runAndWait()
