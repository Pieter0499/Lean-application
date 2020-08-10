import PySimpleGUI as sg
import webbrowser
from ctypes import windll
import tkinter as tk
from tkinter import filedialog
from PIL import ImageGrab


#aan de hand van width en height ratio berekenen
## mogelijk ook achter komen hoe Python kan weten wat de display scale is? Zodat ie, als niet 100%, een waarschuwing kan geven?
user32 = windll.user32
heigth = user32.GetSystemMetrics(1)
width = user32.GetSystemMetrics(0)

print(width)

ratio = 0.8*(heigth/1080)
letter = round(10*ratio)
#Window 0: Intro
introtekst = "Hello user! This application will guide you to a roadmap that will help\n\
make a plan to make the assembly Augmented Reality-ready. This application\n\
is based on the Master thesis by Valentin Strauss named 'Development of a \n\
smart technology roadmap tool towards Augmented Reality based assembly in \n\
Lean enterprises.' To develop this roadmap, different questions will be  \n\
asked, such as to see if the Assembly "


if heigth < 1080:
    waarschuwing = "WARNING: your resolution is lower than 1080p, possibly due to scaling.\nThis will hinder the program. Right click on desktop to adjust display scaling.\nRestart program then."
else:
    waarschuwing = ""

layout_intro = [
        [sg.Text(waarschuwing,text_color="red")],    
        [sg.Text(introtekst)],
        #[sg.Image(r"C:/Users/slmpp/Documents/Ontwikkeling van een LeanAssesment tool/afbeeldingen/bestlogo.png")],
        [sg.Button('Next page',key="-Next1-",pad=((200,0), (30,0)))]
        ]

window_intro = sg.Window("Introduction",layout_intro,size=(500,500))

#Loop en Window 1
while True:
    event, values = window_intro.read()
    if event == sg.WIN_CLOSED:
        break
    if event == "-Next1-":
        layout_2 = [
            [sg.Text('1.1. In the first questionnaire, some general information is asked\n\regarding the company. The questions can be skipped.')],
            [sg.Text('Country of operation:')],[sg.Input(size=(20,0))],
            [sg.Text('Customer base (Local, EU, International):')],[sg.Input(size=(20,0))],
            [sg.Text('Number of employees:')],[sg.Input(size=(20,0))],
            [sg.Text('Share of employees being operators (parts pickers, \nassemblers, machinery operators):')],[sg.Input(size=(20,0))],
            [sg.Text('Rough estimate of annual turnover:')],[sg.Input(size=(20,0))],
            [sg.Text('Share of product costs that assembly represents:')],[sg.Input(size=(20,0))],
            [sg.Text('What type of products are being produced?:')],[sg.Input(size=(20,0))],
            [sg.Text('Position(s) of the responent(s):')],[sg.Input(size=(20,0))],
            [sg.Button('Next page',key='-Next2-',pad=((200,0),(50,0)))]
            ]
        window_1 = sg.Window("1.1. General information",layout_2,size=(500,600),resizable=True)
        break
window_intro.close()

#Loop en Window 2
while True:
    event, values = window_1.read()
    if event == sg.WIN_CLOSED:
        break
    if event == '-Next2-':
        layout_3 = [
            [sg.Text('1.2. Now, some questions regarding knowledge and capabilites\n\about Industry 4.0 are asked. These can also be skipped.'),sg.Button('?',key='-Info-')],
            [sg.Text('Is management aware of the potential performance\nbenefits that Industry 4.0 technologies offer?:')],[sg.Input(size=(20,0))],
            [sg.Text('Is management considering implementation of Industry 4.0\ntechnologies and concepts in their business?')],[sg.Input(size=(20,0))],
            [sg.Text('Has the company conducted experiments to study the\nusability of AR in its assembly context?')],[sg.Input(size=(20,0))],
            [sg.Text('If yes, what were the outcomes of the experiment?')],[sg.Multiline(autoscroll=True)],
            [sg.Text('If no, has the company experimented with other enhancing\ncognitive aid technologies?')],[sg.Multiline(autoscroll=True)],
            [sg.Button('Next page',key='-Next3-')]
            ]    
        window_2 = sg.Window("1.2. Industry 4.0 knowledge",layout_3,size=(500,500))
        break
window_1.close()

#loop en Window 3
while True:
    event, values = window_2.read()
    if event == sg.WIN_CLOSED:
        break
    if event == '-Info-':
        layout_info = [
            [sg.Text("Industry 4.0 is the term used to describe the new industrial revolution\n\
                currently taking place in manufacturing. This concept aims at making use\n\
                of new technologies (Internet of Things, Cloud Computing, Big Data and Analytics,\n\
                Simulation, Augmented Reality, Additive Manufacturing, Collaborative Robots)\n\
                to integrate, digitalize and automate every part of a company. By harvesting\n\
                performance data from the company's various assessts and further analyze them,\n\
                it is possible to create 'virtual twin factory' from which can be simulated a widerange\n\
                of scenarios. This in turn provides the company with a self-optimization capabilities,\n\
                making it's operations able to automatically adapt. Industry 4.0 also has a human-centric\n\
                approach which aim at augmenting and integrating the employees with the technologies\n\
                surrounding them. The outcomes of going towards Industry 4.0 are increased adaptability,\n\
                increased flexibility of production (batch size, variants management) and efficiency.")]
            ]
        window_info = sg.Window("Info",layout_info,grab_anywhere=True)
        event, values = window_info.read()
        window_info.close()
    if event == '-Next3-':
        layout_4 = [
            [sg.Text("1.3. The last introduction questionnaire will be about the Company's\nexperience with Lean. Again, they may be skipped."),sg.Button('?',key='-Info1-')],
            [sg.Text('Has there been information sharing throughout the company to increase\nawareness on the benefits of Lean practices?')],[sg.Input(size=(20,0))],
            [sg.Text('Which Lean practices have already or are being implemented in the company?')],[sg.Multiline(autoscroll=True)],
            [sg.Text('Has the company and its employees adopted a continuous improvement\nmindset aimed at constantly improving efficiency, costs and quality?')],[sg.Input(size=(20,0))],
            [sg.Text('Are the performances after implementation satisfactory for management?')],[sg.Input(size=(20,0))],
            [sg.Button('Next page',key='-Next4-')]]
        window_3 = sg.Window("1.3. Lean knowledge",layout_4,size=(500,500))
        break   
window_2.close()

#loop en Window 4
while True:
    event, values = window_3.read()
    if event == sg.WIN_CLOSED:
        break
    if event == '-Next4-':
        layout_5 = [
            [sg.Text('2.1. The following questionnaires are regarding the assembly.\nThey help to see which\nsteps have to be\ntaken to make the assembly Augmented Reality-ready.')],
            [sg.Text('What is your current position in the market?')],[sg.Input(size=(20,0))],
            [sg.Text('Are there desires from management to evolve the position?')],[sg.Input(size=(20,0))],
            [sg.Text('What are the goals of your strategy over the next 5 years?')],[sg.Multiline(autoscroll=True)],
            [sg.Text('How is your current strategy going to change?')],[sg.Multiline(autoscroll=True)],            
            [sg.Text('What role has assembly in this strategy?')],[sg.Multiline(autoscroll=True)],            
            [sg.Text('In this strategy, what performances need to be improved\nin assembly? (Quality, lead-time, costs, safety, adaptability,…)')],[sg.Multiline(autoscroll=True)],
            [sg.Text('Have you identified the assembly activities and\ntheir respective complexities to overcome in order to achieve\nthe needed perfromances improvements?'),sg.Button("?",key='-Info-')],[sg.Combo(['Yes','No'],key='-Assembly-',default_value='Yes',readonly=True,size=(10,0))],
            [sg.Button('Next page',key='-Next5-',pad=((200,0),0))]
            ]
        window_4 = sg.Window('2.1. Assembly information',layout_5,size=(500,700),resizable=True)
        break
window_3.close()

#loop en Window 5
while True:
    event, values = window_4.read()
    if event == sg.WIN_CLOSED:
        break
    if event == '-Info-':
        #layout_info = [
        #    [sg.Text('Verdere uitleg')],
        #    [sg.Image(r"C:/Users/slmpp/Documents/Ontwikkeling van een LeanAssesment tool/afbeeldingen/Complexities.png")]
        #    ]
        #titel = 'Complexities information'
        #indow_info = sg.Window(titel,layout_info)
        #event, values = window_info.read()
        #window_info.close()
        webbrowser.open_new('https://www.youtube.com/watch?v=1BSIznVzy3o')
    if event == '-Next5-':
        if values['-Assembly-'] == 'Yes':
            layout_6 = [
                [sg.Text('As you have answered yes, some additional questions are presented:')],
                [sg.Text("Are there activities which complexities would be overcome by\ndisplaying through AR real-time instructions?\n\nIf so, click on the button below to answer\nthe questionnaire regarding this."),sg.Button('Why AR?',key='1')],[sg.Button('Answer questionnaire',key='guidance')],
                [sg.Text('\nIs performance lacking because training is required often\nand is time demanding?\n\nIf so, click on the button below to answer\nthe questionnaire regarding this.'),sg.Button('Why AR?',key='2')],[sg.Button('Answer questionnaire',key='training')],
                [sg.Text('\nIs performance lacking in part fetching due to the variety of\nparts and components to be gathered in various locations?\n\nIf so, click on the button below to answer\nthe questionnaire regarding this.'),sg.Button('Why AR?',key='3')],[sg.Button('Answer questionnaire',key='partspicking')]
                ]
        if values['-Assembly-'] == 'No':
            layout_6 = [
                [sg.Text('As you have answered no, the following questions will\nhelp identify Augmented Reality opportunities')],
                [sg.Text('Are there activities which mean of improvement would\nbe a real-time display of information (instructions, drawings)?\n\nIf so, click on the button below to answer\nthe questionnaire regarding this.'),sg.Button('Why AR?',key='4')],[sg.Button('Answer questionnaire')],
                [sg.Text('Is the experience of operators impacting their\nperformances when assembling a new product variant?\n\nIf so, click on the button below to answer\nthe questionnaire regarding this.'),sg.Button('Why AR?',key='5')],[sg.Button('Answer questionnaire')],
                [sg.Text('Part of the assembly time is dependant on the\nfetching of parts to the assembly stations?\n\nIf so, click on the button below to answer\nthe questionnaire regarding this.'),sg.Button('Why AR?',key='6')],[sg.Button('Answer questionnaire')]
                ]        
        window_5 = sg.Window('2.2. Assembly questionnaire',layout_6,size=(500,500))
        break
window_4.close()

#loop en window 6
ARhelp = {'1':"What Augmented Reality can do to help:\n\nAsssembly AR real-time instructions guidance -->\nIncrease in throughput, decrease error rate",
'2':"What Augmented Reality can do to help:\n\nAR for operators training -->\nDecrease in training time, Creates interchangeability\nfor multifunctional teams, independent training",
'3':"What Augmented Reality can do to help:\n\nAR guidance for parts picking in warehouse -->\nShortens collectiong times, decrease picking errors",
'4':"What Augmented Reality can do to help:\n\nAsssembly AR real-time instructions guidance -->\nIncrease in throughput, decrease error rate",
'5':"What Augmented Reality can do to help:\n\nAR for operators training -->\nDecrease in training time, Creates interchangeability\nfor multifunctional teams, independent training",
'6':"What Augmented Reality can do to help:\n\nAR guidance for parts picking in warehouse -->\nShortens collectiong times, decrease picking errors"
}

#voor questionnaires
questionnaire = {
'guidance':[[sg.Button('Hide'),sg.Button('See/Update roadmap',pad=((300,0),0),key='-Guidance-')],[sg.Text("This questionnaire will seek what Lean/Industry 4.0\ncapabilities are missing or not regarding AR real-time assembly instructions guidance.\n\nAnswer the questions related to Lean capabilities:")],[sg.Column([
    [sg.Text("Assembly processing sequence is optimized to attain\nmaximum flow and throughput time:")],[sg.Radio('Yes', "RADIO1",key='4 1'),sg.Radio('No', "RADIO1",default=True)],
    [sg.Text("Operators are inclined to follow the instruction sequence:")],[sg.Radio('Yes', "RADIO2",key='4 2'),sg.Radio('No', "RADIO2", default=True)],
    [sg.Text("Instructions, assembly aids and training are formalized:")],[sg.Radio('Yes', "RADIO3",key='3 1'),sg.Radio('No', "RADIO3", default=True)],
    [sg.Text("Company ensures operators have ergonomic work conditions:")],[sg.Radio('Yes', "RADIO4",key='1 1'),sg.Radio('No', "RADIO4", default=True)],
    [sg.Text("Employees are trained for safe operations:")],[sg.Radio('Yes', "RADIO5",key='1 2'),sg.Radio('No', "RADIO5", default=True)],
    [sg.Text("5S principles are applied and respected on the assembly\nlocation (Sort, Set in order, Shine, Standardize, Sustain):")],[sg.Radio('Yes', "RADIO6",key='1 3'),sg.Radio('No', "RADIO6", default=True)],
    [sg.Text("Production management and Team leaders are involved\nin the planning and sequencing of assembly:")],[sg.Radio('Yes', "RADIO7",key='1 4'),sg.Radio('No', "RADIO7", default=True)],
    [sg.Text("Assembly processes are standardized throughout the\nvariants of products:")],[sg.Radio('Yes', "RADIO8",key='2 1'),sg.Radio('No', "RADIO8", default=True)],
    [sg.Text("Management involves and consults all stakeholders\nwhen implementing changes in operations:")],[sg.Radio('Yes', "RADIO9",key='0 1'),sg.Radio('No', "RADIO9", default=True)],
    [sg.Text("Management makes efforts to convince employees to welcome change:")],[sg.Radio('Yes', "RADIO10",key='0 2'),sg.Radio('No', "RADIO10", default=True)]
        ],scrollable=True,vertical_scroll_only=True,size=(500,200))],
    [sg.Text("Answer the questions related to Industry 4.0 capabilities:")],[sg.Column([
    [sg.Text("An ERP system has been implemented for for ressources\nand orders management:")],[sg.Radio('Yes', "RADIO11",key='5 1'),sg.Radio('No', "RADIO11", default=True)],
    [sg.Text("Scheduling of assembly activities through the ERP system:")],[sg.Radio('Yes', "RADIO12",key='5 2'),sg.Radio('No', "RADIO12", default=True)],
    [sg.Text("All employees input and visualize their performances\nand resources (used and needed) in the ERP system:")],[sg.Radio('Yes', "RADIO13",key='5 3'),sg.Radio('No', "RADIO13", default=True)],
    [sg.Text("Assembly instructions are digitalized (not only\ntexts but also drawings):")],[sg.Radio('Yes', "RADIO14",key='7 1'),sg.Radio('No', "RADIO14", default=True)],
    [sg.Text("Development of new products is done using a CAD software:")],[sg.Radio('Yes', "RADIO15",key='7 2'),sg.Radio('No', "RADIO15", default=True)],
    [sg.Text("For parts coming from suppliers, are you also supplied\nwith drawings of said parts:")],[sg.Radio('Yes', "RADIO16",key='6 1'),sg.Radio('No', "RADIO16", default=True)],
    [sg.Text("Parts and components drawings are stored in a PDM/PLM\nsystem:")],[sg.Radio('Yes', "RADIO17",key='6 2'),sg.Radio('No', "RADIO17", default=True)],
    [sg.Text("PDM and PLM data are integrated to the ERP software:")],[sg.Radio('Yes', "RADIO18",key='8 1'),sg.Radio('No', "RADIO18", default=True)],
    [sg.Text("Information and Data from all information systems are\naccessible throughout the company departments:")],[sg.Radio('Yes', "RADIO19",key='8 2'),sg.Radio('No', "RADIO19", default=True)],
    [sg.Text("Assembly data (completion times, errors, wrong parts,\nquality deffects, etc) is acquired and stored:")],[sg.Radio('Yes', "RADIO20",key='9 1'),sg.Radio('No', "RADIO20", default=True)],
    [sg.Text("Acquired Data is processed and analyzed to identify\nimprovement areas:")],[sg.Radio('Yes', "RADIO21",key='9 2'),sg.Radio('No', "RADIO21", default=True)],
    [sg.Text("Operators and other employees are trained to use new,\nmore efficient technologies:")],[sg.Radio('Yes', "RADIO22",key='10 1'),sg.Radio('No', "RADIO22", default=True)],
    ],scrollable=True,vertical_scroll_only=True,size=(500,250))]
    ],
'training':[[sg.Button('Hide'),sg.Button('See/Update roadmap',pad=((300,0),0),key='-Training-')],[sg.Text("This questionnaire will seek what Lean/Industry 4.0\ncapabilities are missing or not regarding\nAR for operators training.\n\nAnswer the questions related to Lean capabilities:")],[sg.Column([
    [sg.Text("Assembly processing sequence is optimized  to attain\nmaximum flow and throughput time:")],[sg.Radio('Yes', "RADIO23",key='4 1'),sg.Radio('No', "RADIO23", default=True)],
    [sg.Text("Operators are inclined to follow the instruction sequence:")],[sg.Radio('Yes', "RADIO24",key='4 2'),sg.Radio('No', "RADIO24", default=True)],
    [sg.Text("Instructions, assembly aids and training are formalized:")],[sg.Radio('Yes', "RADIO25",key='3 1'),sg.Radio('No', "RADIO25", default=True)],
    [sg.Text("Company ensures operators have ergonomic work conditions:")],[sg.Radio('Yes', "RADIO26",key='1 1'),sg.Radio('No', "RADIO26", default=True)],
    [sg.Text("Employees are trained for safe operations:")],[sg.Radio('Yes', "RADIO27",key='1 2'),sg.Radio('No', "RADIO27", default=True)],
    [sg.Text("5S principles are applied and respected on the assembly\nlocation (Sort, Set in order, Shine, Standardize, Sustain):")],[sg.Radio('Yes', "RADIO28",key='1 3'),sg.Radio('No', "RADIO28", default=True)],
    [sg.Text("Production  management and Team leaders are involved in\nthe planing and sequencing of assembly:")],[sg.Radio('Yes', "RADIO29",key='1 4'),sg.Radio('No', "RADIO29", default=True)],
    [sg.Text("Assembly processes are standardized throughout the variants\nof products:")],[sg.Radio('Yes', "RADIO30",key='2 1'),sg.Radio('No', "RADIO30", default=True)],
    [sg.Text("Management involves and consults all stakeholders when\nimplementing changes in operations:")],[sg.Radio('Yes', "RADIO31",key='0 1'),sg.Radio('No', "RADIO31", default=True)],
    [sg.Text("Management makes efforts to convince employees to welcome change:")],[sg.Radio('Yes', "RADIO32",key='0 2'),sg.Radio('No', "RADIO32", default=True)]]
    ,scrollable=True,vertical_scroll_only=True,size=(500,200))],
    [sg.Text("Answer the questions related to Industry 4.0 capabilities:")],[sg.Column([
    [sg.Text("An ERP system has been implemented for for ressources\nand orders management:")],[sg.Radio('Yes', "RADIO33",key='5 1'),sg.Radio('No', "RADIO33", default=True)],
    [sg.Text("Scheduling of assembly activities through the ERP system:")],[sg.Radio('Yes', "RADIO34",key='5 2'),sg.Radio('No', "RADIO34", default=True)],
    [sg.Text("All employees input and visualize their performances\nand resources (used and needed) in the ERP system:")],[sg.Radio('Yes', "RADIO35",key='5 3'),sg.Radio('No', "RADIO35", default=True)],
    [sg.Text("Assembly instructions are digitalized (not only texts\nbut also drawings):")],[sg.Radio('Yes', "RADIO36",key='7 1'),sg.Radio('No', "RADIO36", default=True)],
    [sg.Text("Development of new products is done using a CAD software:")],[sg.Radio('Yes', "RADIO37",key='7 2'),sg.Radio('No', "RADIO37", default=True)],
    [sg.Text("For parts coming from suppliers, are you also supplied with drawings of said parts:")],[sg.Radio('Yes', "RADIO38",key='6 1'),sg.Radio('No', "RADIO38", default=True)],
    [sg.Text("Parts and components drawings are stored in a PDM/PLM system:")],[sg.Radio('Yes', "RADIO39",key='6 2'),sg.Radio('No', "RADIO39", default=True)],
    [sg.Text("PDM and PLM data are integrated to the ERP software:")],[sg.Radio('Yes', "RADIO40",key='8 1'),sg.Radio('No', "RADIO40", default=True)],
    [sg.Text("Information and Data from all information systems are\naccessible throughout the company departments:")],[sg.Radio('Yes', "RADIO41",key='8 2'),sg.Radio('No', "RADIO41", default=True)],
    [sg.Text("Assembly data (completion times, errors, wrong parts,\nquality deffects, etc) is acquired and stored:")],[sg.Radio('Yes', "RADIO42",key='9 1'),sg.Radio('No', "RADIO42", default=True)],
    [sg.Text("Acquired Data is processed and analyzed to identify\nimprovement areas:")],[sg.Radio('Yes', "RADIO43",key='9 2'),sg.Radio('No', "RADIO43", default=True)],
    [sg.Text("Operators and other employees are trained to use new,\nmore efficient technologies:")],[sg.Radio('Yes', "RADIO44",key='10 1'),sg.Radio('No', "RADIO44", default=True)]
    ],scrollable=True,vertical_scroll_only=True,size=(500,250))]
    ],
'partspicking':[[sg.Button('Hide'),sg.Button('See/Update roadmap',pad=((300,0),0),key='-Partspicking-')],[sg.Text("This questionnaire will seek what Lean/Industry 4.0\ncapabilities are missing or not regarding\nAR for parts picking.\n\nAnswer the questions related to Lean capabilities:")],[sg.Column([
    [sg.Text("Assembly processing sequence is optimized  to attain\nmaximum flow and throughput time:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Operators are inclined to follow the instruction sequence:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Instructions, assembly aids and training are formalized:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Company ensures operators have ergonomic work conditions:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Employees are trained for safe operations:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("5S principles are applied and respected on the assembly\nlocation and in the parts storage location. (Sort,\nSet in order, Shine, Standardize, Sustain):")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Production  management and Team leaders are involved in\nthe planing and sequencing of assembly:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Management involves and consults all stakeholders when\nimplementing change:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Management makes efforts to convince employees to welcome change:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))]
    ],scrollable=True,vertical_scroll_only=True,size=(500,200))],
    [sg.Text("Answer the questions related to Industry 4.0 capabilities:")],[sg.Column([
    [sg.Text("An ERP system has been implemented for for ressources\nand orders management:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Scheduling of assembly activities through the ERP system:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Inventory of parts is digitalized, including location\nand quantities of stored parts and components:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("All employees input and visualize their performances\nand resources (used and needed) in the ERP system:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("WMS and/or MES is integrated to the ERP system:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Information and Data from all information systems are\naccessible throughout the company departments:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Operations departments have access to real time\nperformance data from assembly:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Terminals to follow flow of parts and components\nare available in the warehouse:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Assembly data (completion times, errors, wrong parts,\nquality deffects, etc) is acquired and stored:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Acquired Data is processed and analyzed to identify\nimprovement areas:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))],
    [sg.Text("Operators and other employees are trained to use new,\nmore efficient technologies:")],[sg.Combo(['Yes','n/a','No'],default_value='Yes',readonly=True,size=(10,0))]
    ],scrollable=True,vertical_scroll_only=True,size=(500,250))]
    ]
}

questionnaire_titel = {'guidance':'Questionnaire regarding Guidance AR','training':'Questionnaire regarding Training AR','partspicking':'Questionnaire regarding Parts Picking AR'}
questionnaire_window = {'guidance':0,'training':0,'partspicking':0}
questionnaire_aan ={'guidance':0,'training':0,'partspicking':0}
questionnaire_EvVal = {'guidance':[0,0],'training':[0,0],'partspicking':[0,0]}

#functies die de vragen koppelen met de stapjes in de roadmap
def Convert(string): 
    li = list(string.split(" ")) 
    return li 

def Link(d0,stapjes):
    temp = {}
    d1 = {}
    for i in d0:
        if isinstance(i,str) == False:
            continue
        a = Convert(i)
        if a[0] in d1:
            d1[a[0]].append(d0[i])
        else:
            d1[a[0]] = [d0[i]]
    for j in stapjes:
        if stapjes[j][2] not in d1 or False in d1[stapjes[j][2]]:
            temp[j] = stapjes[j]
            continue
        #else:
        #    if False in d1[stapjes[j][2]]:
        #        temp[j] = stapjes[j]
    return temp


#voor roadmaps
#[rij,column,stap_id,pijl_id...] = [x,y,stap_id,pijl_id] ## PLAATS PER RIJ. 
#voor de stap_id wordt de stap met een vraag gekoppeld. Elke stap heeft een unieke stap_id (string!!).
# Echter, vul bij de stapjes die op x-positie 0 zitten als stap_id '-1' in
# De stap_id wordt gekoppeld met de juiste vragen door te kijken in de juiste stapjes dictionary wat de key is van de vraag 
#Voor de pijl_id: kijk hoeveel "doorgetrokken" pijlen er zijn. 1 pijl kan twee stappen verbinden,
# maar kan ook doorgaan na de tweede stap, en een derde, vierde, vijfde etc. stap verbinden. 
# Een pijl_id slaat dus op een unieke pijl, die niet per se slechts enkel twee stappen verbindt.  

#G T roadmap
stapjes_g = {"Continous improvement/Company\nand employees attitude\ntowards change":[0,0,'-1'],
"Human Resource Management":[0,1,'-1'],"Implement 5S organization\nfollowingoperators preferences.":[1,1,'1',"pijl1","pijl2"],
"Standardization":[0,2,'-1'],"Standardization of\nassembly processes.":[3,2,'2',"pijl1","pijl3"],
"Visual Management":[0,3,'-1'],"Formalize assembly\ninstructions and training.":[3,3,'3',"pijl1"],
"JIT/Flow Pull":[0,4,'-1'],"Optimization of\nassembly sequence.":[3,4,'4',"pijl2","pijl3"],
"ICT Implementation\nand Vertical Integration":[0,5,'-1'],"Implement ERP software\nfor resources management\nand activities scheduling":[1,5,'5',"pijl4"],
"Digital Product Development":[0,6,'-1'],"Storing of parts\ndrawings and data\nusing PDM/PLM":[1,6,'6',"pijl5"],"Digitalization of\nassembly instructions":[4,6,'7',"pijl1","pijl6"],
"Horizontal integration of\nsocio-technical systems":[0,7,'-1'],"Integration of PLM\nand PDM systems\nwith ERP":[2,7,'8',"pijl4","pijl5"],
"Data aqcuisition\nand processing.":[0,8,'-1'],"Collection and input\nof performances data\nin ERP":[2,8,'9',"pijl4"],
"ICT and Smart technology\nemployee training":[0,9,'-1'],"Training of operators\nto use the system":[9,9,'10',"pijl8","pijl9"],
"Identify activities and\ntheir complexities needing\nperformances improvement\nthrough AR.":[0,10,'-1'],"Identify assembly activities\nwith AR potential":[3,10,'11'],
"Define displayed content.":[0,11,'-1'],"Involve stakeholders*\nto define the\nneeded content to\nbe displayed":[4,11,'12',"pijl1","pijl7"],
"Choose hardware:":[0,12,'-1'],"Choose hardware.":[5,12,'13',"pijl7"],
"Develop software with\ntracking and content":[0,13,'-1'],"Development of the\nsoftware for tracking\nand display of\ncontent with subcontractor":[6,13,'14',"pijl1","pijl6"],
"Start implementing/\nusing AR":[0,14,'-1'],"First pilot":[7,14,'15',"pijl1","pijl8"],"Adapt system based\non feedback":[9,14,'16',"pijl1"],"Larger scale implementation":[11,14,'17',"pijl1","pijl9"],
"Evaluate new performances\nand Improve":[0,15,'-1'],"Evaluate new assembly\nperformances, continously\nimprove and update":[12,15,'18',"pijl1","pijl4"]
}

stapjes_t = stapjes_g

stapjes_pp = {"Continous improvement/\nCompany and\nemployees attitude towards change":[0,0,'-1'],
"Human Resource Management":[0,1,'-1'],"Organization of\nWarehouse and\nassembly stations\nfollowing 5S":[1,1,'1',"stap1","stap2"],
"Visual management":[0,2,'-1'],"Formalization of\ninstructions and\ntraining.":[2,2,'2',"stap1"],
"JIT/Flow pull":[0,3,'-1'],"Optimization of\nparts collection\npath":[4,3,'3',"stap2"],
"ICT Implementation\nand Vertical Integration":[0,4,'-1'],"ERP software\nfor resources\nand orders.\nScheduling of\nassembly activites\nusing ERP.":[1,4,'4',"stap3","stap4","stap5"],"Digitalization of the warehouse parts\ninventory using Warehouse management system":[3,4,'5',"stap2"],
"Horizontal integration\nof socio-technical\nsystems":[0,5,'-1'],"MES* and/or\nWMS* system integrated\nto ERP":[5,5,'6',"stap2","stap3"],
"Human to\nMachine Interface":[0,6,'-1'],"Optional: integrate\nAR with\nERP for\nfaster order\nactivity intake\nor choose other\ninterface for link\nto ERP/MES/WMS":[7,6,'7',"stap3","stap4"],
"Data acquisition\nand processing\n(Only for AR monitoring)":[0,7,'-1'],"Collection and\ninput of\nperformances data\nin ERP":[4,7,'8',"stap5"],
"ICT and Smart\ntechnology employee\ntraining":[0,8,'-1'],"Training of\noperators to\nuse the system":[10,8,'9',"stap8","stap9"],
"Define displayed content.":[0,9,'-1'],"Define with stakeholders\nwhat information\nto display.":[5,9,'10',"stap6","stap7"],
"Choose hardware.":[0,10,'-1'],"Choose hardware":[6,10,'11',"stap7"],
"Software development":[0,11,'-1'],"Development of\ntracking and\ndisplay software\nwith subcontractors":[7,11,'12',"stap6"],
"Start implementing/using AR":[0,12,'-1'],"First pilot experiment.":[8,12,'13',"stap6","stap8"],"Adapt system\nbased on feedback":[10,12,'14',"stap6"],"Implement/use":[12,12,'15',"stap6","stap9"],
"Evaluate new\nperformances and Improve":[0,13,'-1'],"Evaluate performances\nand continously improve":[13,13,'16',"stap5","stap6"]
}

def Stap(teken,tekst,px,py,kleur='black'):
    lengte = []
    b = 0
    for i in tekst:
        b = b + 1
        if i == '\n':
            lengte.append(b)
            b = -1
    lengte.append(b)
    max_x = max(lengte)
    max_y = len(lengte)
    tloc = sg.TEXT_LOCATION_TOP_LEFT
    if px != 0:
        teken.DrawRectangle(top_left=(px,py+max_y*10*ratio),bottom_right=(px+max_x*7*ratio,py-max_y*10*ratio),fill_color='Orange',line_color='Black',line_width=5*ratio)
        tloc = "center"
        kleur = 'black'
    else:
        max_x = 0
        py = py-max_y*10*ratio
    
    teken.DrawText(tekst,location=(px+0.5*max_x*7*ratio,py),color=kleur,font=(None,letter),angle=0,text_location=tloc)

def Pijl(teken,x0,x1,y0,y1,kolom_d,rij_d):
    px0 = 0
    if x0 > 0:
        for j in range(x0):
            if j not in kolom_d:
                continue
            px0 = px0 + kolom_d[j]*7*ratio + 20
        px0 = px0 + 20
    px1 = 0
    if x1 > 0:
        for j in range(x1):
            if j not in kolom_d:
                continue
            px1 = px1 + kolom_d[j]*7*ratio + 20
        px1 = px1 + 20

    #y0 > n, waarbij n het laagste y-coordinaat is.
    # ook moet het zijn 'for i in range(n+1,y0+1)'
    py0 = 20
    if y0 > 0:
        for i in range(1,y0+1):
            py0 = py0 + 20 + rij_d[i-1]*10*ratio + (rij_d[i]-1)*10*ratio
    py1 = 20
    if y1 > 0:
        for i in range(1,y1+1):
            py1 = py1 + 20 + rij_d[i-1]*10*ratio + (rij_d[i]-1)*10*ratio   
    
    if y1 != y0:
        if abs(y1-y0) > abs(x1-x0):
            teken.DrawLine(point_from=(px0,py0),point_to=(px1+kolom_d[x1]*3.5*ratio,py0),color="black",width=2)
            teken.DrawLine(point_from=(px1+kolom_d[x1]*3.5*ratio,py0),point_to=(px1+kolom_d[x1]*3.5*ratio,py1),color="black",width=2) 
        else:
            teken.DrawLine(point_from=(px0+kolom_d[x0]*3.5*ratio,py0),point_to=(px0+kolom_d[x0]*3.5*ratio,py1),color="black",width=2)
            teken.DrawLine(point_from=(px0+kolom_d[x0]*3.5*ratio,py1),point_to=(px1,py1),color="black",width=2)
    else:
        teken.DrawLine(point_from=(px0,py0),point_to=(px1,py1),color="black",width=2)

#MAAK HIERONDER EEN DEFINITIE VAN!!! goed?
#check hoe "dik" maximaal elke rij is, en hoe "breed" elke kolom is

def functionRoadmap(stapjes,rm_window):
    dik_subsub = []
    rij_dict = {}
    kolom_dict = {}

    for i in stapjes:
        b = 0 
        for j in i:
            b = b + 1
            if j == '\n':
                dik_subsub.append(b-1)
                b = 0
        dik_subsub.append(b)

        rij = stapjes[i][1]

        if rij in rij_dict:
            rij_dict[rij].append(len(dik_subsub))
        else:
            rij_dict[rij] = [len(dik_subsub)]

        kolom = stapjes[i][0]
        if kolom in kolom_dict:
            kolom_dict[kolom].append(max(dik_subsub))
        else:
            kolom_dict[kolom] = [max(dik_subsub)]
        
        dik_subsub = []

    graph_x = 20
    graph_y = 20

    for i in kolom_dict:
        kolom_dict[i] = max(kolom_dict[i])
        graph_x = graph_x + kolom_dict[i]*8 + 20
    for i in rij_dict:
        rij_dict[i] = max(rij_dict[i])
        graph_y = graph_y + rij_dict[i]*25

#niet draggable window, column weg
    layout = [[sg.Button("Return to update roadmap",key='-Return-'),sg.Button("Save as image",pad=((500,0),0),key='-Save-')],[sg.Column([[sg.Graph((graph_x+100,graph_y),graph_bottom_left=(-100,graph_y),graph_top_right=(graph_x,0),key='graph')]],scrollable=False,size=(graph_x+50,graph_y))]]
    rm_window = sg.Window("lol",layout,no_titlebar=True,resizable=True,size=(width,heigth)).Finalize()
    rm_window.Maximize()
    graph = rm_window['graph'] 

    print(graph_x+100)

    y_raster = 20
    vorige_raster = 0
    begin = 0
    uitleg_dic = {"Lean":[1,2,3,4],"I4.0":[5,6,7,8,9],"Implementation":[10,11,12,13,14,15]}

    a = 0
    y_oud = 0
    for i in rij_dict:
        if begin == 0:
            graph.DrawLine(point_from=(0,y_raster-rij_dict[i]*10*ratio-5),point_to=(graph_x*ratio,y_raster-rij_dict[i]*10*ratio-5),color="white",width=0.5)
            vorige_raster = rij_dict[i]
            begin = 1
        else:
            #prev_y = y_raster
            y_raster = y_raster + vorige_raster*10*ratio + (rij_dict[i]-1)*10*ratio + 20
            #graph.DrawRectangle(top_left=(0,prev_y-vorige_raster*10-5),bottom_right=(graph_x,y_raster-rij_dict[i]*10-5),fill_color=kleur2,line_color=kleur2,line_width=5)
            graph.DrawLine(point_from=(0,y_raster-rij_dict[i]*10*ratio-5),point_to=(graph_x*ratio,y_raster-rij_dict[i]*10*ratio-5),color="white",width=0.5)
            vorige_raster = rij_dict[i]
        
        for k in uitleg_dic:
            if a == 1:
                graph.DrawRectangle(top_left=(-90,y_oud),bottom_right=(-20,y_raster-rij_dict[i]*10-5),fill_color="grey",line_color="black",line_width=5)
                graph.DrawText(tekst,location=(-55,(y_oud+(y_raster-rij_dict[i]*10*ratio-5))/2),color='black',font=None,angle=90,text_location="center")
                y_oud = y_raster-rij_dict[i]*10*ratio-5
                a = 0
            tekst = k
            if i == uitleg_dic[k][-1]:
                a = 1
                break
    graph.DrawRectangle(top_left=(-90,y_oud),bottom_right=(-20,graph_y),fill_color="grey",line_color="black",line_width=5)
    graph.DrawText(tekst,location=(-55,(y_oud+(y_raster-rij_dict[i]*10*ratio-5))/2),color='black',font=None,angle=90,text_location="center")

    
    graph.DrawLine(point_from=(kolom_dict[0]*7*ratio+20,0),point_to=(kolom_dict[0]*7*ratio+20,graph_y),color="white",width=0.5)
#y = y + 20 + rij_dict[stapjes[i][1]-1]*10 + (rij_dict[stapjes[i][1]]-1)*10

    y = 20
    x = 0
    vorige = 0

    for i in stapjes:
        pijl_n = len(stapjes[i])-3
        if pijl_n < 1:
            continue
        for j in range(pijl_n):
            x0_ = stapjes[i][0]
            y0_ = stapjes[i][1]
            nu_zoeken = 0
            pijl_id = stapjes[i][j+3]
            for k in stapjes:
                if k == i:
                    nu_zoeken = 1
                if nu_zoeken == 1 and k != i:
                    if pijl_id in stapjes[k]:
                        x1_ = stapjes[k][0]
                        y1_ = stapjes[k][1]
                        Pijl(teken=graph,x0=x0_,x1=x1_,y0=y0_,y1=y1_,kolom_d = kolom_dict,rij_d = rij_dict)
                        break
    graph.DrawRectangle(top_left=(kolom_dict[0]*7*ratio+40,20-rij_dict[0]*7*ratio),bottom_right=(graph_x*ratio,20+rij_dict[0]*7*ratio),fill_color='Orange',line_color='Black',line_width=5*ratio)
    graph.DrawText("Involvement of stakeholders in discussions, before and during the project to gather opinions, doubts, and ideas. Efforts to convince operators of change necessity.",location=((kolom_dict[0]*7*ratio+40+graph_x*ratio)/2,20),color='black',font=(None,letter),angle=0,text_location="center")
    
    for i in stapjes:
        x = 0
        if stapjes[i][0] > 0:
            for j in range(stapjes[i][0]):
                if j not in kolom_dict:
                    continue
                x = x + kolom_dict[j]*7*ratio+20
            x = x + 20
        if stapjes[i][1] != vorige:
            y = y + 20 + rij_dict[stapjes[i][1]-1]*10*ratio + (rij_dict[stapjes[i][1]]-1)*10*ratio
        Stap(teken=graph,tekst=i,px=x,py=y)
        vorige = stapjes[i][1]
    
    while True:
        ev, val = rm_window.read()
        if ev == '-Return-':
            break
        if ev == '-Save-':
            try:
                #myScreenshot = pyautogui.screenshot()
                myScreenshot = ImageGrab.grab()
                file_path = filedialog.asksaveasfilename(defaultextension='.png')
                myScreenshot.save(file_path)
            except:
                continue
    rm_window.close()

roadmap_stapjes = {'-Guidance-':stapjes_g,
'-Training-':stapjes_t,
'-Partspicking-':stapjes_pp
}

roadmap_window = {'-Guidance-':0,'-Training-':0,'-Partspicking-':0}

#roadmap_EvVal = 

#NU DOEN DAT IE UPDATE IPV EEN NIEUWE WINDOW STARTEN, DAARNA SCREENSHOT??
#OOK PARTSPICKING DOEN, NIEUWE STAPJES MAKEN
while True:
    event, values = window_5.read(timeout=100)
    for i in questionnaire_aan:
        if questionnaire_aan[i] == 1:
            questionnaire_EvVal[i,0], questionnaire_EvVal[i,1] = questionnaire_window[i].read(timeout=100)
            if questionnaire_EvVal[i,0] == 'Hide':
                questionnaire_window[i].hide()
            if questionnaire_EvVal[i,0] in roadmap_window:
                c = Link(d0 = questionnaire_EvVal[i,1],stapjes=roadmap_stapjes[questionnaire_EvVal[i,0]])
                functionRoadmap(stapjes=c,rm_window=roadmap_window[questionnaire_EvVal[i,0]])
    if event in ARhelp:
        layout_info = [[sg.Text(ARhelp[event])]]
        titel = 'What Augmented Reality can help to do'
        window_info = sg.Window(titel,layout_info)
        event, values = window_info.read()
        window_info.close()
    if event in questionnaire:
        if questionnaire_aan[event] == 0:
            titel = questionnaire_titel[event]
            questionnaire_window[event] = sg.Window(titel,questionnaire[event],size=(500,500),grab_anywhere=True,no_titlebar=True)
            questionnaire_aan[event] = 1
        else:
            questionnaire_window[event].UnHide()
       


window_5.close()
