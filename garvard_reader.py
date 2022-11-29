import pandas as pd
from requests_html import HTMLSession
import pyttsx3
from random import shuffle
import os
from os.path import abspath
from gtts import gTTS 
from playsound import playsound 

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[2].id)
engine.setProperty('rate', 90)
engine.runAndWait()

path = abspath(__file__)
index = path.rfind('\\')
path = path[:index]
def whatTake(text,libRed,lang='en'):
    if libRed == '1':
        speakGtts(text,lang)
    elif libRed =='2':
        speak(text)
def speakGtts(text,lang='en'):
    var = gTTS(text = text,lang = lang) 
    var.save(f'{desktop}/eng.mp3')
    playsound(f'{desktop}/eng.mp3')
    os.remove(f'{desktop}/eng.mp3')
def speak(str):
    engine.say(str)
    engine.runAndWait()


def takeSomelements(list_, count):
    while True:
        try:
            return list_[:count]
        except:
            count -= 1
            continue


words = pd.read_excel(f'{desktop}\\words.xlsx')['WORDS']
libRed = input('1.GTTS \n 2.pyttsx3\n')
start = input('Start point\n')
if start == '':
    start = 0
else:
    start = int(start) - 2

if len(words) < start or start < 0:
    raise Exception

end = input('End point\n')
if end == '':
    end = len(words)
else:
    end = int(end) - 2

if len(words) < end or end < 0:
    raise Exception

rand = input('Random? \n Y/N')
matrixForWords = [_ for _ in range(start, end + 1)]
print()
if rand.lower() == 'y':
    shuffle(matrixForWords)

for posWord in matrixForWords:
    print('-' * 30)
    print(words[posWord], '   ', posWord + 2, end='\n\n')
    count = 0
    whatTake(f'Word . {words[posWord]} . in number {posWord+2}',libRed)
    url = f'https://dictionary.cambridge.org/ru/словарь/англо-русский/{words[posWord]}'
    session = HTMLSession()
    while True:
        
        try:
        
            r = session.get(url)
            break
        except:
            continue
    for body in r.html.xpath('//div[@class="entry-body"]'):

        thisWord = body.xpath('//span[@class="hw dhw"]')[0].text
        try:
            part = body.xpath(
                '//span[@class="pos dpos"]')[0].text + body.xpath(
                    '//span[@class="spellvar dspellvar"]')[0].text
        except:
            try:
                
                part = body.xpath('//span[@class="pos dpos"]')[0].text
            except:
                part=''
        if part:
            
            print(f'{thisWord} as {part}')
            whatTake(f'Word . {thisWord} . as {part}',libRed)
            
        for block in body.xpath(
                '//div[(@class="sense-body dsense_b" and (./div[@class="def-block ddef_block"])) or @class="pr phrase-block dphrase-block"]'
        ):
            for meaning, examples in zip(
                    block.xpath('//div[@class="def ddef_d db"]'),
                    block.xpath('//div[@class="def-body ddef_b"]')):
                count += 1
                whatTake(f'{count} meaning',libRed)
                phrase = block.xpath(
                    '//span[@class="phrase-title dphrase-title"]')
                if phrase:
                    print(f'Phrase - {phrase[0].text}')
                    whatTake(f'It"s phrase . {phrase[0].text.replace("/"," or ")}',libRed)

                rusMe = examples.xpath('//span')[0].text

                example = examples.xpath('//div[@class="examp dexamp"]')
                example = '.\n'.join([
                    f"· {example_.text.strip('.')}"
                    for example_ in takeSomelements(example, 10)
                ]).replace('/', ' or ').replace('\\', ' or ') + '.'
                print(meaning.text)
                print(rusMe)
                print(f'{example}', end='\n\n')
                whatTake(meaning.text,libRed)
                engine.setProperty('voice', voices[0].id)
                engine.setProperty('rate', 150)
                whatTake(rusMe,libRed,'ru')
                engine.setProperty('voice', voices[2].id)
                whatTake('EXAMPLES',libRed)
                engine.setProperty('rate', 90)
                whatTake(example,libRed)
