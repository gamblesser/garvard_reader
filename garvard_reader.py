import pandas as pd
from requests_html import HTMLSession
import pyttsx3

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[2].id)
engine.setProperty('rate', 90)
engine.runAndWait()

def speak(str):
    engine.say(str)
    engine.runAndWait()

def takeSomelements(list_,count):
    while True:
        try:    
            return list_[:count]
        except:
            count-=1
            continue

words = pd.read_excel('1000 most used english words.xlsx')['WORDS']
start = int(input('Start point\n'))
for k,wordForSearch in enumerate(words[start-2:]):
    print('-'*30)
    print(wordForSearch,'   ',start+k,end='\n\n')
    count = 0
    speak(f'Word . {wordForSearch} . in number {start+k}')
    url = f'https://dictionary.cambridge.org/ru/словарь/англо-русский/{wordForSearch}'
    session = HTMLSession()
    r = session.get(url)
    for meaning,examples in zip(r.html.xpath('//div[@class="def ddef_d db"]'),r.html.xpath('//div[@class="def-body ddef_b"]')):
        count+=1
        speak(f'{count} meaning')
        
        
        rusMe= examples.xpath('//span')[0].text
        
        example= examples.xpath('//div')[1:]
        example= '.\n'.join([f"· {example_.text.strip('.')}" for example_ in takeSomelements(example,10)]).replace('/',' ').replace('\\',' ')+'.'
        print(meaning.text)
        print(rusMe)
        print(f'{example}',end='\n\n')
        
        
        speak(meaning.text)
        engine.setProperty('voice', voices[0].id)
        engine.setProperty('rate', 130)
        tts = speak(rusMe)
        engine.setProperty('voice', voices[2].id)
        speak('EXAMPLES')
        engine.setProperty('rate', 90)
        tts = speak(example)
