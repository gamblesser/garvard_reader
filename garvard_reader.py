import pandas as pd
from requests_html import HTMLSession
import pyttsx3
from random import shuffle

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[2].id)
engine.setProperty('rate', 90)
engine.runAndWait()


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


words = pd.read_excel('1000 most used english words.xlsx')['WORDS']

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
    speak(f'Word . {words[posWord]} . in number {posWord+2}')
    url = f'https://dictionary.cambridge.org/ru/словарь/англо-русский/{words[posWord]}'
    session = HTMLSession()
    r = session.get(url)
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
            speak(f'Word . {thisWord} . as {part}')
            
        for block in body.xpath(
                '//div[(@class="sense-body dsense_b" and (./div[@class="def-block ddef_block"])) or @class="pr phrase-block dphrase-block"]'
        ):
            for meaning, examples in zip(
                    block.xpath('//div[@class="def ddef_d db"]'),
                    block.xpath('//div[@class="def-body ddef_b"]')):
                count += 1
                speak(f'{count} meaning')
                phrase = block.xpath(
                    '//span[@class="phrase-title dphrase-title"]')
                if phrase:
                    print(f'Phrase - {phrase[0].text}')
                    speak(
                        f'It"s phrase . {phrase[0].text.replace("/"," or ")}')

                rusMe = examples.xpath('//span')[0].text

                example = examples.xpath('//div[@class="examp dexamp"]')
                example = '.\n'.join([
                    f"· {example_.text.strip('.')}"
                    for example_ in takeSomelements(example, 10)
                ]).replace('/', ' or ').replace('\\', ' or ') + '.'
                print(meaning.text)
                print(rusMe)
                print(f'{example}', end='\n\n')

                speak(meaning.text)
                engine.setProperty('voice', voices[0].id)
                engine.setProperty('rate', 150)
                tts = speak(rusMe)
                engine.setProperty('voice', voices[2].id)
                speak('EXAMPLES')
                engine.setProperty('rate', 90)
                tts = speak(example)
