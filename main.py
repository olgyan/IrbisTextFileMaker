"""
This program makes an IRBIS entry (in txt format for importing) from an
ISBD-style bibliographical reference, to save time cataloguing publications.
Written in Python 3.8 because my work computer only has Windows 7
and it was not possible to install later versions of Python.
Olga Yanovskaya
"""

import os
import re
from datetime import date
from platform import system
from tkinter import Tk, Text, NW, LEFT, StringVar, TclError
from tkinter.ttk import Frame, Button, Label
from tkinter.filedialog import askopenfilename, asksaveasfilename

import config
import rusmarc

regex_ozboz = re.compile(r'.+?(?= [/:;]|$)', re.IGNORECASE)
regex_heading = re.compile(r'^([\w]+?),? ((?:[\w{1,2}]\. ){1,2})(.+)')
regex_publ = re.compile(r'([^:]*?)(?: ?: (.*?))?, (\d{4})')
regex_m_pages = re.compile(r'^(\d+) ([сcpSsл]\.?)')
regex_s_pages = re.compile(r'^([CСPS]\.) ?(\d+\D{1,3}\d+)')
regex_edition = re.compile(r'^\d{1,2}(?=-е из).+')
regex_issue = re.compile(
    r'''((?:issue|no\.|v(?:\.|ol(?:\.|ume))|вып(?:\.|уск)|
выпуск №|т(?:\.|ом\.?)|ч(?:\.|асть)|№){1,2} ?[-(),/\d IVXDC]+)''',
    re.IGNORECASE)

if system() == 'Linux':
    app_dir = config.linux_app_dir
    irbiswrk_dir = config.linux_irbiswrk_dir
elif system() == 'Windows':
    app_dir = config.windows_app_dir
    irbiswrk_dir = config.windows_irbiswrk_dir
else:
    app_dir = ''
    irbiswrk_dir = ''
today = date.today().strftime('%Y%m%d')
dashes = str.maketrans('–—―‒−', '-----')
noyo = str.maketrans('Ёё', 'Ее')
irbis_entry = {}
parsing_batch = []
place = ()


class Field:
    """Irbis field
    attributes:
        tag - field tag
        contents - a list of dictionaries:
            occ. no. is the index ([0] is all occs. together)
            subfield tag is the key of the dictionary
            contents of that subfield are the corresponding value
            subfield tag is '_' if the field has no subfields
            subf_list == contents[0]
    """

    def __init__(self, tag: int):
        """Makes an empty field object."""
        self.tag = tag
        self.contents = [{}]
        self.subf_list = self.contents[0]
        irbis_entry[self.tag] = self

    def fill(self, value):
        """Fills fields with values.
        Updates subfield value lists in self.contents[0]."""
        global place
        v_occ, v_subf = place[1:3]
        o = len(self.contents)
        while o < v_occ + 1:
            self.contents.append({})
            o = len(self.contents)
        if value:
            self.contents[v_occ][v_subf] = value
        kk = set()
        for o in self.contents:
            kk.update(o.keys())
        for key in kk:
            self.subf_list[key] = []
            for i, o in enumerate(self.contents):
                if i != 0:
                    self.subf_list[key].append(o.get(key))

    def show(self, pft=None):
        """Turns field object to string for printing or writing to file.
        Field -> generator of strings, one for every occ.
        print(*field.show('allpft'), sep='\n')

        'pft' attribute defines printing format:
            'allpft' - imitates all.pft
            'irbistext' - makes text file for importing
        """
        if pft == 'irbistext':
            tag_form = f'#{self.tag}: '
        else:
            tag_form = ''
        for occ_no, occ_text in zip(range(1, len(self.contents)),
                                    self.contents[1:]):
            if pft == 'allpft':
                tag_form = f'#{self.tag}/{occ_no}:_'
            if '_' in occ_text:
                yield f'{tag_form}{str(occ_text["_"])}'
            else:
                fitext = ''.join(f'^{k}{v}' for k, v in occ_text.items())
                yield f'{tag_form}{fitext}'


class Parsing:

    def __init__(self, bibref):

        bibref = bibref.strip(' \r\n\t')
        while '  ' in bibref:
            bibref = bibref.replace('  ', ' ')
        bibref = bibref.translate(dashes)
        bibref = bibref.replace('.- ', '. - ')
        bibref = bibref.translate(noyo)

        single_work, _, source = bibref.partition(r' // ')
        self.asp = bool(source)
        ozboz, *desc_areas = single_work.split('. - ')
        title, title_info, resp_stmt_1, resp_stmt_2 = self.make_title(ozboz)
        fld('200^a').fill(title)
        fld('200^e').fill(title_info)
        fld('200^f').fill(resp_stmt_1)
        fld('200^g').fill(resp_stmt_2)
        self.make_author(701, resp_stmt_1)
        if resp_stmt_2:
            self.make_author(702, resp_stmt_2)

        if self.asp:
            s_ozboz, *larger_coll = source.split('. - ')
            (s_title, s_title_info, s_resp_stmt_1,
             s_resp_stmt_2) = self.make_title(s_ozboz)
            if s_resp_stmt_1:
                self.make_author(961, s_resp_stmt_1)
            if s_resp_stmt_2:
                self.make_author(961, s_resp_stmt_2)
            fld('463^c').fill(s_title)
            fld('963^e').fill(s_title_info)
            if 'конференц' in s_title_info:
                fld('900^c').fill('18f')
            if s_resp_stmt_1:
                if s_resp_stmt_2:
                    fld('963^f').fill(f'{s_resp_stmt_1} ; {s_resp_stmt_2}')
                else:
                    fld('963^f').fill(s_resp_stmt_1)
            desc_areas += larger_coll

        for area in desc_areas:
            publ = re.search(regex_publ, area)
            if type(publ) is re.Match:
                publ_info = publ.groups()
                fno, s = (463, 'dgj') if self.asp else (210, 'acd')
                no_publ = ('б. м.', 'б. и.', 'б. г.')
                for j in range(3):
                    fld(f'{fno}^{s[j]}').fill(
                        publ_info[j] if publ_info[j] else no_publ[j])

            year = re.match(r'\d{4}', area)
            if type(year) is re.Match:
                fld('463^j').fill(year[0])
                fld('900^b').fill('08')
            elif self.asp:
                fld('900^b').fill('09')
            elif not self.asp:
                fld('900^b').fill('05')

            pages_regex = regex_s_pages if self.asp else regex_m_pages
            pages = re.search(pages_regex, area)
            if type(pages) is re.Match:
                fno, st = (463, '1s') if self.asp else (215, 'a1')
                for j in range(2):
                    fld(f'{fno}^{st[j]}').fill(pages[j + 1])

            ed_area = re.search(regex_edition, area)
            if type(ed_area) is re.Match:
                ed_info = ed_area.group().split(', ')
                for s, val in zip('ab', ed_info):
                    fld(f'205^{s}').fill(val)

            issue = [
                elem.rstrip(', ') for elem in re.findall(regex_issue, area)
            ]
            for s, val in zip('vil', issue):
                fld(f'463^{s}').fill(val)

            if area.startswith('EDN'):
                for s, val in zip(
                        'tih',
                    ('Ссылка на публикацию',
                     f'https://elibrary.ru/item.asp?edn={area[4:10]}', '05')):
                    fld(f'951^{s}').fill(val)

            if area.startswith('ISBN'):
                code = ('961^i') if self.asp else ('10^a')
                fld(code).fill(area[5:])

            if area.startswith('DOI'):
                fld('19^A').fill('6 DOI')
                fld('19^B').fill(area[4:])

        fld('920').fill('ASP' if self.asp else 'PAZK')
        for s, val in zip('cab', ('ПК', today, 'itfmaker')):
            fld(f'907^{s}').fill(val)

        self.fields = irbis_entry.copy()
        irbis_entry.clear()
        parsing_batch.append(self)

        print(f'*****\n\n{bibref}\n\n')
        for field in self.fields.values():
            print(*field.show('allpft'), sep='\n')

    def make_title(self, area):
        ozboz = re.findall(regex_ozboz, area)
        title, title_info, resp_stmt_1, resp_stmt_2, heading = (
            '' for x in range(5))
        if z := re.match(regex_heading, ozboz[0]):
            title = z[3]
            fld('700^a').fill(z[1])
            fld('700^b').fill(z[2].strip())
        else:
            title = ozboz[0]
        for elem in ozboz[1:]:
            if elem.startswith(' : '):
                title_info = elem[3:].replace(elem[3], elem[3].lower())
                continue
            elif elem.startswith(' / '):
                resp_stmt_1 = elem[3:]
                continue
            elif elem.startswith(' ; '):
                resp_stmt_2 = elem[3:]
                continue
        return title, title_info, resp_stmt_1, resp_stmt_2

    def make_author(self, fno, resp_stmt):
        if '[и др.]' in resp_stmt:
            resp_stmt = resp_stmt[:-8]
        people = resp_stmt.split(', ')
        for i, person in enumerate(people):
            if not bool(re.search('[\u0400-\u04FF]', person)):
                notify_text.set('Имена латиницей, нужно будет исправить.')

            nameparts = person.split()
            famname = ' '.join(
                [n for n in nameparts if '.' not in n and n[0].isupper()])
            io = ' '.join(
                [n for n in nameparts if '.' in n and n[0].isupper()])
            ed_func = ' '.join(
                [n for n in nameparts if n not in famname and n[0].islower()])

            fld(f'{fno}^a#{i+1}').fill(famname)
            fld(f'{fno}^b#{i+1}').fill(io)
            if ed_func:
                fld(f'{fno}^4#{i+1}').fill(
                    str(rusmarc.fcode.get(ed_func, 570)) + ed_func)
            if 700 in irbis_entry and 701 in irbis_entry:
                try:
                    irbis_entry[701].contents.remove(
                        irbis_entry[700].contents[1])
                except ValueError:
                    pass

    def choose_mono_genre(self, title_info):
        mono_genre = {
            'монография': '22',
            'научно-популярная литература': '19',
            'учебник': 'j0',
            'учебное пособие': 'j'
        }

        if title_info in mono_genre:
            fld('900^c').fill(mono_genre[title_info])


def fld(code):
    """get or create a field object by an IRBIS-like designation
    code is TAG^SUBF#OCC, like &uf('av900^A#1') in Irbis
    place stores position in fields to communicate it between functions
    #n makes new occurrences (does not work well: makes a new occ for every
    subfield, use for fields with no subfields only, like 610)
    """
    global irbis_entry, place
    if '^' in code:
        v_tag = code[:code.find('^')]
    elif '#' in code:
        v_tag = code[:code.find('#')]
    else:
        v_tag = code
    v_tag = int(v_tag)
    field = irbis_entry[v_tag] if v_tag in irbis_entry else Field(v_tag)
    v_subf = code[code.index('^') + 1].upper() if '^' in code else '_'
    v_occ = code[code.index('#') + 1:] if '#' in code else 1
    if v_occ == 'n':
        v_occ = len(field.contents)
    place = v_tag, int(v_occ), v_subf
    return field


def paste(form):
    """check if it is an ISBD ref. that we are pasting in the form"""
    try:
        clipboard = form.clipboard_get()
        clipboard = clipboard.translate(dashes)
        if '. - ' in clipboard:
            form.delete('1.0', 'end')
            form.insert('1.0', clipboard)
        else:
            notify_text.set('В буфере обмена не библиографическая ссылка.')
    except TclError:
        notify_text.set('В буфере обмена ничего не было.')


def do_parsing():
    """get a single bib.ref. from entry form and parse"""
    app.config(cursor='watch')
    try:
        notify_text.set('Разбираем...')
        paste(isbd_get_text)
        Parsing(isbd_get_text.get('1.0', 'end'))
        notify_text.set('Разобрано')
    except IndexError:
        notify_text.set('Нечего разбирать.')
    except Exception as e:
        notify_text.set(f'Что-то пошло не так: {e}')
    finally:
        app.config(cursor='arrow')


def multi_parsing():
    """get many bib.refs. from text file and parse"""
    filepath = askopenfilename(title='Выбор файла со ссылками',
                               initialdir=f'{app_dir}test_values',
                               defaultextension='txt')
    if filepath != '':
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            bibrefs = f.readlines()
            for line in bibrefs:
                Parsing(line)


def save_irbis_text():
    """ ask for a file to save parsings to and save them"""
    filepath = asksaveasfilename(
        title='Сохранить текстовый файл для импорта в ИРБИС',
        initialdir='C:\\irbiswrk',
        defaultextension='txt',
        initialfile=f'import_{today}.txt')
    if filepath != '':
        with open(filepath,
                  'w',
                  encoding='utf-8',
                  errors='ignore',
                  newline='\r\n') as f:
            for parsing in parsing_batch:
                for field in parsing.fields.values():
                    f.writelines(f'{line}\n'
                                 for line in field.show('irbistext')
                                 if len(line) > 6)
                f.write('*****\n')
        app.destroy()


def readme():
    os.system(f'{app_dir}readme.txt')


app = Tk()
app.title('Разбираем ссылку для импорта в ИРБИС')

up = Frame(app)
paste_btn = Button(up,
                   text='Вставить',
                   command=lambda isbd_get_text: paste(isbd_get_text))
workapart_btn = Button(up, text='Разобрать', command=do_parsing)
fromfile_btn = Button(up, text='Разобрать из файла', command=multi_parsing)
save_btn = Button(up,
                  text='Сохранить в текст. файл ИРБИС',
                  command=save_irbis_text)
readme_btn = Button(up, text='readme', command=readme)
isbd_get_text = Text(app, wrap='word', height=5)
down = Frame(app)
notify_text = StringVar()
notify = Label(down, textvariable=notify_text)

up.pack(pady=2, anchor=NW)
paste_btn.pack(padx=2, side=LEFT)
workapart_btn.pack(padx=2, side=LEFT)
fromfile_btn.pack(padx=2, side=LEFT)
save_btn.pack(padx=2, pady=1, side=LEFT)
readme_btn.pack(padx=2, pady=1, side=LEFT)
isbd_get_text.pack(padx=2, pady=1, anchor=NW)
down.pack(pady=1, anchor=NW)
notify.pack(padx=0, side=LEFT)
app.mainloop()
