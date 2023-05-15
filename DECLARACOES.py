import PySimpleGUI as sg
from docxtpl import DocxTemplate
import subprocess
sg.theme('Green')

#layout1 menu de declarações
layout1 = [
            [sg.Text("Tipo da declaração:")],
            [sg.Combo(['DECLARAÇÃO-GERAL-ADRIANA','VINCULO-ESTUDANTIL','DECLARAÇÃO-INSTRUMENTAÇÃO','DECLARAÇÃO-ESTÁGIOS-SUPERVISIONADOS-TARDE'],key='tipo',size=(50,0))],
            [sg.Button('PROXIMO')]
        ]

#layout2 DECLARAÇÃO-GERAL-ADRIANA
layout2 = [
            [sg.Text('Nome:',size=(5,0)),sg.Input(size=(29,0),key='nome',do_not_clear=False)],
            [sg.Text('RG:',size=(5,0)),sg.Input(size=(10,0),key='rg',do_not_clear=False),sg.Text('EXP:',size=(4,0)),sg.Input(size=(10,0),key='rgtipo',do_not_clear=False)],
            [sg.Text('CPF:',size=(5,0)),sg.Input(size=(29,0),key='cpf',do_not_clear=False)],
            [sg.Text("Curso:",size=(5,0)),sg.Combo(['TÉCNICO EM ENFERMAGEM','TÉCNICO EM ANÁLISES CLÍNICAS','TÉCNICO EM RADIOLOGIA', 'TÉCNICO EM SEGURANÇA DO TRABALHO'],key='curso',size=(29,0))],
            [sg.Text('Hora:',size=(5,0)),sg.Input(size=(5,0),key='hora1',do_not_clear=False),sg.Text('ás',size=(2,0)),sg.Input(size=(5,0),key='hora2',do_not_clear=False)],
            [sg.Text('Data(ex:janeiro de 2023):',size=(18,0)),sg.Input(size=(15,0),key='data',do_not_clear=False)],
            [sg.Text('Turno:',size=(5,0)),sg.Combo(['manhã','tarde','noite'],key='turno')],
        ]
#layout3 VINCULO-ESTUDANTIL
layout3 = [
            [sg.Text('Nome:',size=(5,0)),sg.Input(size=(29,0),key='nome1',do_not_clear=False)],
            [sg.Text('RG:',size=(5,0)),sg.Input(size=(10,0),key='rg1',do_not_clear=False),sg.Text('EXP:',size=(4,0)),sg.Input(size=(10,0),key='rgtipo1',do_not_clear=False)],
            [sg.Text('CPF:',size=(5,0)),sg.Input(size=(29,0),key='cpf1',do_not_clear=False)],
            [sg.Text("Curso:",size=(5,0)),sg.Combo(['TÉCNICO EM ENFERMAGEM','TÉCNICO EM ANÁLISES CLÍNICAS','TÉCNICO EM RADIOLOGIA', 'TÉCNICO EM SEGURANÇA DO TRABALHO'],key='curso1',size=(29,0))],
            [sg.Text('Hora:',size=(5,0)),sg.Input(size=(5,0),key='hora11',do_not_clear=False),sg.Text('ás',size=(2,0)),sg.Input(size=(5,0),key='hora21',do_not_clear=False)],
            [sg.Text('Inicio das aulas para o mês(ex:janeiro):',size=(15,0)),sg.Input(size=(18,0),key='mes',do_not_clear=False)],
            [sg.Text('Ano(ex:2023):',size=(10,0)),sg.Input(size=(24,0),key='ano',do_not_clear=False)],
            [sg.Text('Data(ex:janeiro de 2023):',size=(18,0)),sg.Input(size=(15,0),key='data1',do_not_clear=False)],
            [sg.Text('Turno:',size=(5,0)),sg.Combo(['manhã','tarde','noite'],key='turno1')],
        ]
#layout4 DECLARAÇÃO-INSTRUMENTAÇÃO
layout4 = [
            [sg.Text('Nome:',size=(5,0)),sg.Input(size=(29,0),key='nome2',do_not_clear=False)],
            [sg.Text('RG:',size=(5,0)),sg.Input(size=(10,0),key='rg2',do_not_clear=False),sg.Text('EXP:',size=(4,0)),sg.Input(size=(10,0),key='rgtipo2',do_not_clear=False)],
            [sg.Text('CPF:',size=(5,0)),sg.Input(size=(29,0),key='cpf2',do_not_clear=False)],
            [sg.Text('Curso:',size=(5,0)),sg.Combo(['TÉCNICO EM ENFERMAGEM','TÉCNICO EM ANÁLISES CLÍNICAS','TÉCNICO EM RADIOLOGIA', 'TÉCNICO EM SEGURANÇA DO TRABALHO','INSTRUMENTAÇÃO CIRÚRGICA'],key='curso2',size=(29,0))],
            [sg.Text('Local(ex:Hospital de Trauma):',size=(23,0)),sg.Input(size=(20,0),key='local',do_not_clear=False)],
            [sg.Text('Hora:',size=(5,0)),sg.Input(size=(5,0),key='hora12',do_not_clear=False),sg.Text('ás',size=(2,0)),sg.Input(size=(5,0),key='hora22',do_not_clear=False)],
            [sg.Text('Data(ex:janeiro de 2023):',size=(18,0)),sg.Input(size=(25,0),key='data2',do_not_clear=False)],
            [sg.Text('Turno:',size=(5,0)),sg.Combo(['manhã','tarde','noite'],key='turno2')],

        ]

layout5 = [
            [sg.Text('Nome:',size=(5,0)),sg.Input(size=(29,0),key='nome3',do_not_clear=False)],
            [sg.Text('RG:',size=(5,0)),sg.Input(size=(10,0),key='rg3',do_not_clear=False),sg.Text('EXP:',size=(4,0)),sg.Input(size=(10,0),key='rgtipo3',do_not_clear=False)],
            [sg.Text('CPF:',size=(5,0)),sg.Input(size=(29,0),key='cpf3',do_not_clear=False)],
            [sg.Text('Curso:',size=(5,0)),sg.Combo(['TÉCNICO EM ENFERMAGEM','TÉCNICO EM ANÁLISES CLÍNICAS','TÉCNICO EM RADIOLOGIA', 'TÉCNICO EM SEGURANÇA DO TRABALHO','INSTRUMENTAÇÃO CIRÚRGICA'],key='curso3',size=(29,0))],
            [sg.Text('Local(ex:Hospital de Trauma):',size=(23,0)),sg.Input(size=(20,0),key='local2',do_not_clear=False)],
            [sg.Text('Tipo de estagio:',size=(12,0)),sg.Input(size=(33,0),key='tipoestagio',do_not_clear=False)],
            [sg.Text('No periodo de:',size=(10,0)),sg.Input(size=(5,0),key='periodo1',do_not_clear=False),sg.Text('a',size=(1,0)),sg.Input(size=(10,0),key='periodo2',do_not_clear=False)],
            [sg.Text('Data(ex:janeiro de 2023):',size=(18,0)),sg.Input(size=(25,0),key='data3',do_not_clear=False)],
            [sg.Text('Turno:',size=(5,0)),sg.Combo(['manhã','tarde','noite'],key='turno3')],

        ]


#Menu
layout = [[sg.Column(layout1, key='-COL1-'), sg.Column(layout2, visible=False, key='-COL2-'), sg.Column(layout3, visible=False, key='-COL3-'),sg.Column(layout4, visible=False, key='-COL4-'),sg.Column(layout5, visible=False, key='-COL5-')],
          [sg.Button('Exit'),sg.Button('Voltar'),sg.Button('GERAR DECLARAÇÃO')]]

window = sg.Window('Gerador de declação FESVIP', layout,icon='vista-panoramica.ico', element_justification='c')

layout = 1  #Primeiro layout que vai aparecer
while True:
    event, values = window.read()
    tipo = values['tipo']
    print(event, values)
    if event in (None, 'Exit'):
        break
    elif event in 'PROXIMO':
        if tipo == 'DECLARAÇÃO-GERAL-ADRIANA':
            loc = ''
            layout = 1
            layout = layout + 1
            window[f'-COL{layout}-'].update(visible=True)
            window['-COL1-'].update(visible=False)  
        elif tipo == 'VINCULO-ESTUDANTIL':
            loc = 1
            layout = 1
            layout = layout + 2
            window[f'-COL{layout}-'].update(visible=True)
            window['-COL1-'].update(visible=False) 
        elif tipo == 'DECLARAÇÃO-INSTRUMENTAÇÃO':
            loc = 2
            layout = 1
            layout = layout + 3           
            window[f'-COL{layout}-'].update(visible=True)
            window['-COL1-'].update(visible=False)
        elif tipo == 'DECLARAÇÃO-ESTÁGIOS-SUPERVISIONADOS-TARDE':
            loc = 3
            layout = 1
            layout = layout + 4           
            window[f'-COL{layout}-'].update(visible=True)
            window['-COL1-'].update(visible=False)     
    elif event == 'GERAR DECLARAÇÃO' :
        tipo = values['tipo']
        nome = values[f'nome{loc}']
        rg = values[f'rg{loc}']
        rgtipo = values[f'rgtipo{loc}']
        cpf = values[f'cpf{loc}']
        turno = values[f'turno{loc}']
        curso = values[f'curso{loc}']
        if loc == '' or loc < 3:
            hora1 = values[f'hora1{loc}']
            hora2 = values[f'hora2{loc}']
            hora = f"{hora1}hrs às {hora2}hrs"
        else:
            hora = ''
        data = values[f'data{loc}']
        #layout4
        local = values['local']
        ano = values['ano']
        mes = values['mes']
        #layout5
        periodo1 = values['periodo1']
        periodo2 = values['periodo2']
        periodo = f"{periodo1} a {periodo2}"
        local2 = values['local2']
        tipoestagio = values['tipoestagio']
        nome = nome.upper()
        rgtipo = rgtipo.upper()
        curso = curso.upper()

        
        doc = DocxTemplate(f"{tipo}.docx")
        context = { 'nome' : nome , 'rg': rg,'rgtipo':rgtipo, 'cpf': cpf, 'curso': curso, 'turno':turno, 'hora':hora, 'data':data, 'local':local, 'periodo' :periodo,'tipoestagio':tipoestagio, 'local2':local2, 'mes':mes, 'ano':ano}
        #deixar o nome sem espaço
        nomes = nome
        alunos = [aluno.strip().split(' ')[0] for aluno in nomes]
        alunoss = "".join(alunos)
        doc.render(context)
        doc.save(f"{alunoss}{tipo}edit.docx")
        command = f"start {alunoss}{tipo}edit.docx"
        print(subprocess.getoutput(command))      
    elif event in 'Voltar':
        layout = 1
        #quando add mais layouts aumente o range
        for i in range(2, 6):
            window[f'-COL{i}-'].update(visible=False)
        window['-COL1-'].update(visible=True)       

window.close()