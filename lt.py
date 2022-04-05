from datetime import date
import subprocess
import shutil
import json
import click
import docx
import os
import sys
import os.path
import time
import pywinauto
from openpyxl import load_workbook

dt = date.today()
rasdial = subprocess.check_output('rasdial').decode('utf-8')  # for checking VPN connection
CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help', '--h'])

# reads config.json
with open('config.json', 'r') as jsonfile:
    config = json.load(jsonfile)
    jsonfile.close()



@click.group(context_settings=CONTEXT_SETTINGS)
def cli():
    pass


# change prefix (document name before turn number)
@cli.command(name='pf', help='Change the value of \'prefix\' (the beginning of a turn name).')
@click.argument('name')
def pf(name):
    previousname = config['prefix']
    config['prefix'] = name
    with open('config.json', 'w') as jsonfile:
        json.dump(config, jsonfile)
    click.echo(f'\nChanged \'{previousname}\' to \'{name}\'.')


# change speaker names
@cli.command(name='s1', help='Change the value of \'speaker1\'.')
@click.argument('name')
def s1(name):
    previousname = config['speaker1']
    config['speaker1'] = name
    with open('config.json', 'w') as jsonfile:
        json.dump(config, jsonfile)
    click.echo(f'\nChanged \'{previousname}\' to \'{name}\'.')


@cli.command(name='s2', help='Change the value of \'speaker2\'.')
@click.argument('name')
def s2(name):
    previousname = config['speaker2']
    config['speaker2'] = name
    with open('config.json', 'w') as jsonfile:
        json.dump(config, jsonfile)
    click.echo(f'\nChanged \'{previousname}\' to \'{name}\'.')


@cli.command(name='s3', help='Change the value of \'speaker3\'.')
@click.argument('name')
def s3(name):
    previousname = config['speaker3']
    config['speaker3'] = name
    with open('config.json', 'w') as jsonfile:
        json.dump(config, jsonfile)
    click.echo(f'\nChanged \'{previousname}\' to \'{name}\'.')


@cli.command(name='s4', help='Change the value of \'speaker4\'.')
@click.argument('name')
def s4(name):
    previousname = config['speaker4']
    config['speaker4'] = name
    with open('config.json', 'w') as jsonfile:
        json.dump(config, jsonfile)
    click.echo(f'\nChanged \'{previousname}\' to \'{name}\'.')


# create and open new doc
@cli.command(name='doc',
             help='Creates and opens a Word document corresponding with turn specified.')
@click.argument('turn')
def doc(turn):
    config['last_turn'] = config['prefix'] + turn.upper()
    doc = docx.Document('C:/Users/LEE/AppData/Roaming/Microsoft/Templates/AGNSW 2021.docx')
    doc._body.clear_content()
    if turn == 'a':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [10.00 - 10.15]')
    elif turn == 'b':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [10.15 - 10.30]')
    elif turn == 'c':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [10.30 - 10.45]')
    elif turn == 'd':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [10.45 - 11.00]')
    elif turn == 'e':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [11.00 - 11.15]')
    elif turn == 'f':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [11.15 - 11.30]')
    elif turn == 'g':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [11.30 - 11.45]')
    elif turn == 'h':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [11.45 - 12.00]')
    elif turn == 'i':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [12.00 - 12.15]')
    elif turn == 'j':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [12.15 - 12.30]')
    elif turn == 'k':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [12.30 - 12.45]')
    elif turn == 'l':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [12.45 - 1.00]')
    elif turn == 'l2':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [1.00 - 1.15]')
    elif turn == 'm':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [2.00 - 2.15]')
    elif turn == 'n':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [2.15 - 2.30]')
    elif turn == 'o':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [2.30 - 2.45]')
    elif turn == 'p':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [2.45 - 3.00]')
    elif turn == 'q':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [3.00 - 3.15]')
    elif turn == 'r':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [3.15 - 3.30]')
    elif turn == 's':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [3:30 - 3.45]')
    elif turn == 't':
        doc.add_paragraph(f'START OF TURN {config["last_turn"]} [3.45 - 4.00]')
    else:
        click.echo('\nInvalid turn.')
        quit()
    # save turn's path to config.json so that it can be copied to VPN drive when finished
    config['last_turn_path'] = f'C:/Users/LEE/Desktop/{config["last_turn"]}.docx'
    with open('config.json', 'w') as jsonfile:
        json.dump(config, jsonfile)
    doc.save(config['last_turn_path'])
    click.echo(f'\nCreated Word document: {config["last_turn"]}.docx')
    # open document
    os.system(f'start C:/Users/LEE/Desktop/{config["last_turn"]}.docx')
    try:
        app = pywinauto.Application().connect(best_match=config['last_turn'], timeout=5).top_window()
    except pywinauto.timings.TimeoutError:
        click.echo(f'\n{config["last_turn"]}.docx did not open.')
    else:
        app.type_keys('{END}')  # moves cursor to end of document


# open daily folders and write path to config.json
@cli.command(name='daily',
             help='Open folders relevant to today\'s date and presiding officer specified. Writes \'daily_path\' to config.json.')
def daily():
    if 'Legal Transcripts' in rasdial:  # checks if connected to VPN
        list = []
        for parent, dirs, files in os.walk(f'X:/{dt.strftime("%Y")}/{dt.strftime("%B")}/{dt.strftime("%d.%m.%y")}/'):
            for dirname in dirs:
                list.append(dirname)
        click.echo()  # blank line - is there a better way to do this?
        for cnt, name in enumerate(list, 1):
            sys.stdout.write('%d. %s\n\r' % (cnt, name))
        choice = int(input('\nSelect daily folder [1-%s]: ' % cnt)) - 1
        click.echo(f'\nOpening \'{list[choice]}\'.')
        config['daily_path'] = f'X:/{dt.strftime("%Y")}/{dt.strftime("%B")}/{dt.strftime("%d.%m.%y")}/{list[choice]}'
        if os.path.exists(config['daily_path']) == True:
            with open('config.json', 'w') as jsonfile:
                json.dump(config, jsonfile)
            subprocess.Popen(['C:/Program Files/GPSoftware/Directory Opus/dopus.exe', config['daily_path']])
            time.sleep(1)  # allows both folders to open in the same window
            subprocess.Popen(['C:/Program Files/GPSoftware/Directory Opus/dopus.exe', f'S:/AGNSW DAILIES/{dt.strftime("%Y%m%d")}'])
        else:
            click.echo('\nFolder doesn\'t exist.')
            quit()
    else:
        click.echo('\nNot connected to \'Legal Transcripts VPN 2\'.')


# connect/disconnect VPN
@cli.command(name='vpn',
             help='Toggles connection to \'Legal Transcipts VPN 2\'.')
def vpn():
    if 'No connections' in rasdial:
        click.echo('\nNot connected.\n\nConnecting...')
        os.system('start rasphone')
        app = pywinauto.Application().connect(title='Network Connections', timeout=3).top_window()
        app.type_keys('{ENTER 2}')
    else:
        click.echo('\nConnected.\n\nDisconnecting...')
        os.system('rasphone -h "Legal Transcripts VPN 2"')


# saves and copies 'last_turn' to 'daily_path' and saves info in excel
@cli.command(name='save',
             help='Saves and closes \'last_turn\', moves it to \'daily_path\', and writes info to Excel invoice.')
def save():
    try:
        app = pywinauto.Application().connect(best_match=config['last_turn'], timeout=2).top_window()
    except pywinauto.timings.TimeoutError:
        click.echo(f'\n{config["last_turn"]}.docx is not open.')
    else:
        app.type_keys('^s')  # saves the document
        app.close()
        time.sleep(0.2)
        doc = docx.Document(config['last_turn_path'])
        doc.add_paragraph()  # adds blank paragraph - probably a better way to do this
        doc.add_paragraph(f'END OF TURN {config["last_turn"]}')
        doc.save(config['last_turn_path'])
        click.echo(f'\nSaved and closed: {config["last_turn"]}.docx')

    word_count = 0
    for para in doc.paragraphs:
        word_count = word_count + len(para.text.split())
    click.echo(f'\nCounted {word_count} words.')

    wb = load_workbook(filename='C:/Users/LEE/Documents/work/Lee Luppi transcription invoice period end 15.04.22.xlsx')

    empty_row = 0
    for row in wb.active.iter_rows(min_row=15, max_row=81, max_col=1):
        for cell in row:
            if cell.value is None:
                empty_row = cell.row
        if empty_row >= 1:
            break

    wb.active.cell(row=empty_row, column=1).value = config['last_turn']
    wb.active.cell(row=empty_row, column=2).value = dt.strftime('%d.%m.%y')
    wb.active.cell(row=empty_row, column=4).value = word_count
    wb.save('C:/Users/LEE/Documents/work/Lee Luppi transcription invoice period end 15.04.22.xlsx')
    click.echo(f'\nCopied \'{config["last_turn"]}\', \'{dt.strftime("%d.%m.%y")}\', and \'{word_count}\' to row {empty_row}.')

    if 'Legal Transcripts' in rasdial:  # checks if connected to VPN
        shutil.move(config['last_turn_path'], config['daily_path'])
        click.echo(f'\n{config["last_turn"]}.docx moved to \'{config["daily_path"]}\'.')
    else:
        click.echo('\nNot connected to VPN.  Could not move turn to daily folder.')


if __name__ == '__main__':
    cli()
