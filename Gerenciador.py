from tkinter import *
from tkinter import ttk, messagebox,filedialog
import os
import Criar
import sys
import xlsxwriter
from PIL import Image
from random import random
from datetime import datetime
from unidecode import unidecode
import threading
import pandas as pd
import qrcode

minhapasta = os.path.dirname(__file__)
Criar.cria()
abr = 0
abr2 = 0
abr3 = 0
abr4 = 0
abr5 = 0
abr6 = 0
abr7 = 0
abrr = 0
abr9 = 0
abr=8
en = 5
cli = 1
ordenarpor = 0
nomedocli= ""
numerodocle = ""
valo = 0
ativades = 0
quantdodes = 0
sempre = 0
vendapo = 0

if os.path.isdir(minhapasta + '\\Meu estoque'):
    pass
else:
    os.mkdir(minhapasta + '\\Meu estoque')

if os.path.isdir(minhapasta + '\\Meu estoque\\QRCode_unico'):
    pass
else:
    os.mkdir(minhapasta + '\\Meu estoque\\QRCode_unico')


al12 = 0
al2 = 0
def abritelaalt(pro , codes,ambia):
    global al2
    global al12
    al12 = 0
    al2 = 0
    telaaltera = Tk()
    telaaltera.title('Alterar produto')
    telaaltera.geometry('750x400')
    telaaltera.resizable(width=0, height=0)
    telaaltera.iconbitmap(minhapasta+"\\update.ico")

    Label(telaaltera, text='Preencha os dados', fg='#000',  font=('arial', 15)).place(x=50, y=10, width=360, height=38)

    Label(telaaltera, text='Nome do produto:' ,fg='#000',  font=('arial', 15)).place(x=10, y=50, width=200, height=38)
    nom = Label(telaaltera, text=pro, fg='#000',  font=('arial', 15),anchor=W)
    nom.place(x=200, y=50, width=400, height=38)
    nome = Entry(telaaltera,width=25,font=('arial', 13))

    a = Label(telaaltera, text='Ambiente do produto:' ,fg='#000',  font=('arial', 15))
    a.place(x=500, y=150, width=200, height=38)
    amb = Label(telaaltera, text=ambia, fg='#000',  font=('arial', 15))
    amb.place(x=500, y=200, width=200, height=38)

    Label(telaaltera, text='Quant. em unidades:', fg='#000',  font=('arial', 15)).place(x=10, y=100, width=200, height=38)
    qtuni = Entry(telaaltera,width=15,font=('arial', 13))
    qtuni.place(x=220, y=100, width=200, height=30)


    Label(telaaltera, text='Valor por unidades R$:', fg='#000',  font=('arial', 15)).place(x=10, y=150, width=200, height=38)
    valoruni = Entry(telaaltera,width=15,font=('arial', 13))
    valoruni.place(x=220, y=150, width=200, height=30)

    Label(telaaltera, text='Valor total:', fg='#000',  font=('arial', 15)).place(x=100, y=200, width=200, height=38)

    valortotal = '0'

    valort = Label(telaaltera, text='R$: '+ valortotal,fg='#000',  font=('arial', 15))
    valort.place(x=100, y=250, width=200, height=38)

    ambien = ttk.Treeview(telaaltera, columns=(
        'nome'), show='headings')
    ambien.column('nome', minwidth=198,
                    width=198, anchor=CENTER, stretch=NO)

    ambien.heading('nome', text='Ambientes', anchor=CENTER)

    ambien.tag_configure('oddrow', background="lightblue")
    ambien.tag_configure('evenrow', background="white")
    ambien.tag_configure('red', background="orange")

    pesquisaambi = Entry(telaaltera, width=25, font=('arial', 13))

    def cadastrar():
        global al12
        global al2
        a=1
        b = 1.5
        r=0
        if al2 != 2:
            ambiente = ambia
        else:
            try:
                itens = ambien.selection()[0]
                valoresaa = ambien.item(itens, 'values')
                ambiente = valoresaa[0]
            except:
                messagebox.showerror(title='ERROR', message='Selecione o ambiente')
                return
        if al12 == 1:
            nomes = nome.get()
        else:
            nomes = pro
        
        qtunis = qtuni.get()
        valores = valoruni.get()
        if nomes == '' or qtunis == ''or valores == '':
            messagebox.showerror(title='ERROR', message='Preencha os dados')
            telaaltera.lift()
        else:
            try:
                cod = (unidecode(nomes) + " R$:"+str(valoruni.get()))
                img = qrcode.make(cod)
                img.save('barcode.jpg')
                valorunis = float(valores.replace(',','.'))
                type(int(qtunis)) == type(a)
                type(float(valorunis)) == type(b)
            except:
                messagebox.showerror(title='ERROR', message='Preencha os dados corretamente')
                telaaltera.lift()

            if  type(int(qtunis)) == type(a) and type(float(valorunis)) == type(b) and nomes != '':
                    r = 1
            if r == 1:
                valortotals = int(qtunis) * float(valorunis)
                me = Criar.alteraproduto(nomes.lower(),int(qtunis),round(float(valorunis),2),round(float(valortotals),2),cod,codes,int(qtunis),ambiente)
                if me != None:
                    messagebox.showerror(title='ERROR', message=me)
                    telaaltera.lift()
                    
                else:
                    Criar.deletaprodutoquasevv(codes)
                    Criar.deletaprodutoquasevv2(codes)
                    buscarproambiente()
                    buscarHome()
                    telaaltera.quit()
                    telaaltera.destroy()
                    
    def enter(event):
        global al2
        cas = pesquisaambi.get()
        try:
            if al2 == 2 and cas != '':
                buscarambientesas()
            else:
                cadastrar()   
        except:
            pass

    def calcular():
        try:
            uni = float(valoruni.get())
            qt = float(qtuni.get())
            valort.configure(text='R$: '+str(uni*qt))
        except:
            pass

    def altno():
        global al12
        nome.place(x=220, y=50, width=200, height=30)
        nome.focus()
        nom.destroy()
        al12 = 1

    def altamb():
        global al2
        amb.destroy()
        a.destroy()
        altambi.destroy()
        pesquisaambi.place(x=500, y=10, width=150, height=31)
        ambien.place(x=510, y=50, width=200, height=250)
        buscaramb.place(x=660, y=10, width=60, height=30)
        buscarambientesas()
        al2 = 2

    def buscarambientesas():
        ambien.column('nome', minwidth=198,width=198, anchor=CENTER, stretch=NO)

        cor = 0
        ambien.delete(*ambien.get_children())
        vquery = 'SELECT * FROM ambiente WHERE nome LIKE '+'"%'+pesquisaambi.get()+'%"'
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            if cor == 0:
                ambien.insert("", "end", values=i, tags=('oddrow',))
                cor = 1
            else:
                ambien.insert("", "end", values=i, tags=('evenrow',))
                cor = 0

    def qui():
        telaaltera.quit()
        telaaltera.destroy()
    nome.focus()
    
    calbt = Button(telaaltera, text='Verificar', command=calcular)
    calbt.configure(bd=1,
                    activebackground="#467", cursor='hand2')
    calbt.place(x=100, y=330, width=130, height=38)

    altnome =Button(telaaltera, text='Alterar nome', command=altno)
    altnome.configure(bd=1,
                    activebackground="#467", cursor='hand2')
    altnome.place(x=350, y=10, width=100, height=30)

    altambi =Button(telaaltera, text='Alterar ambiente', command=altamb)
    altambi.configure(bd=1,
                    activebackground="#467", cursor='hand2')
    altambi.place(x=550, y=10, width=100, height=30)

    cadastrarbt = Button(telaaltera, text='Alterar', command=cadastrar)
    cadastrarbt.configure(bd=1, 
                    activebackground="#467", cursor='hand2')
    cadastrarbt.place(x=550, y=330, width=100, height=38)

    buscaramb = Button(telaaltera, text='Buscar', command=buscarambientesas)
    buscaramb.configure(bd=1, 
                    activebackground="#467", cursor='hand2')
    
    sairbt = Button(telaaltera, text='Sair', command=qui,width=5)
    sairbt.configure(bd=1, 
                    activebackground="#467", cursor='hand2')
    sairbt.place(x=400, y=330, width=50, height=38)

    telaaltera.wm_protocol('WM_DELETE_WINDOW', qui)
    telaaltera.bind(sequence="<Return>", func=enter)
    telaaltera.mainloop()

def telaAlteraProduto(pro,cods,ambi):
    abritelaalt(pro,cods,ambi)

def frameHome():
    global en
    
    home.place(x=0, y=0, width=1187, height=812)
    produtosVendidos.place(x=-1187, y=0, width=950, height=620)
    vendas.place(x=-1187, y=0, width=950, height=620)
    avancar.place(x=-1187, y=0, width=950, height=620)
    impri.place(x=-1187, y=0, width=1187, height=812)
    histori.place(x=-1187, y=0, width=950, height=620)
    ambientes.place(x=-1187, y=0, width=950, height=620)
    pesquisapro.focus()
    pesquisapro.delete(0,END)
    pesquisapreco.delete(0,END)
    mostrarProHome()
    en = 5
    

def frameVendas():
    global en
    global vendapo
    global cli
    vendas.place(x=0, y=0, width=1187, height=812)
    home.place(x=-1187, y=0, width=950, height=620)
    produtosVendidos.place(x=-1187, y=0, width=950, height=620)
    avancar.place(x=-1187, y=0, width=950, height=620)
    impri.place(x=-1187, y=0, width=1187, height=812)
    histori.place(x=-1187, y=0, width=950, height=620)
    ambientes.place(x=-1187, y=0, width=950, height=620)
    if cli == 1:
        nomeCliente.focus()
    else:
        nomeCliente2.focus()

    if vendapo == 1:
        addProduto.place(x=270, y=600, width=125, height=38)
        alteratudovenda.place(x=420, y=600, width=50, height=38)
        codBarra.place(x=-1000, y=250, width=300, height=38)
        labelqtd.place(x=50, y=600, width=75, height=38)
        qtdadd.place(x=125, y=600, width=75, height=38)
        pesquisavenda.place(x=90, y=150, width=200, height=38)
        buscarvenda.place(x=310, y=152, width=63, height=31)
        recarregarvenda.place(x=390, y=150, width=41, height=38)
        produtosparavenda.place(x=50, y=200, width=400, height=370)
        mostrarProvendas()
        labelcod.configure(text='Lista de produtos:')
    else:
        addProduto.place(x=270, y=313, width=125, height=38)
        alteratudovenda.place(x=420, y=313, width=50, height=38)
        codBarra.place(x=125, y=250, width=300, height=38)
        labelqtd.place(x=50, y=313, width=75, height=38)
        qtdadd.place(x=125, y=313, width=75, height=38)
        pesquisavenda.place(x=-1500, y=150, width=200, height=38)
        buscarvenda.place(x=-1500, y=152, width=63, height=31)
        recarregarvenda.place(x=-1500, y=150, width=41, height=38)
        produtosparavenda.place(x=-1500, y=200, width=654, height=575)
        labelcod.configure(text='Código do produto:')

    codBarra.delete(0,END)
    qtdadd.delete(0,END)
    mostrarquaseVend()
    en = 1
    
def frameProdutosVendidos():
    global en
    produtosVendidos.place(x=0, y=0, width=1187, height=812)
    home.place(x=-1187, y=0, width=950, height=620)
    vendas.place(x=-1187, y=0, width=950, height=620)
    avancar.place(x=-1187, y=0, width=950, height=620)
    impri.place(x=-1187, y=0, width=1187, height=812)
    histori.place(x=-1187, y=0, width=950, height=620)
    ambientes.place(x=-1187, y=0, width=950, height=620)
    pesquisaproVendido.focus()
    pesquisaproVendido.delete(0,END)
    mostrarVend()
    en =3
 

def frameAvancar():
    des.delete(0,END)
    vmetodo.set(metodosdepag[0])
    vstatus.set(statuspag[0])
    v.configure(text="")
    global cli
    
    if cli == 1:
        nome = nomeCliente.get()
        cel = celCliente.get()
    else:
        nome = nomeCliente2.get()
        cel = celCliente2.get()

    if cel == '':
        cel = "00000000000"

    if nome == '' or len(cel) < 9:
        messagebox.showerror(title='ERROR', message="Preencha os dados")
    else:
        global en
        global nomedocli
        global numerodocle
        nomedocli = nome
        numerodocle = cel
        avancar.place(x=0, y=0, width=1187, height=812)
        produtosVendidos.place(x=-1187, y=0, width=950, height=620)
        home.place(x=-1187, y=0, width=950, height=620)
        vendas.place(x=-1187, y=0, width=950, height=620)
        impri.place(x=-1187, y=0, width=1187, height=812)
        histori.place(x=-1187, y=0, width=950, height=620)
        ambientes.place(x=-1187, y=0, width=950, height=620)
        threading.Thread(target=som).start()
        en =0
        

def frameListaimprimir():
    global en
    
    impri.place(x=0, y=0, width=1187, height=812)
    produtosVendidos.place(x=-1187, y=0, width=1187, height=812)
    home.place(x=-1187, y=0, width=950, height=620)
    vendas.place(x=-1187, y=0, width=950, height=620)
    avancar.place(x=-1187, y=0, width=950, height=620)
    histori.place(x=-1187, y=0, width=950, height=620)
    ambientes.place(x=-1187, y=0, width=950, height=620)
    pesquisaproImpri.focus()
    pesquisaproImpri.delete(0,END)
    mostrarimpri()
    en =10

def frameHistorico():
    global en
    global abrr 
    global ordenarpor
    
    histori.place(x=0, y=0, width=1187, height=812)
    impri.place(x=-1187, y=0, width=1187, height=812)
    produtosVendidos.place(x=-1187, y=0, width=1187, height=812)
    home.place(x=-1187, y=0, width=950, height=620)
    vendas.place(x=-1187, y=0, width=950, height=620)
    avancar.place(x=-1187, y=0, width=950, height=620)
    ambientes.place(x=-1187, y=0, width=950, height=620)
    mostarOrdemhis()
    pesquisaprohist.focus()
    abrr = 0
    en =20
    
def frameAmbientes():
    global en
    mostarambients()
    ambientes.place(x=0, y=0, width=1187, height=812)
    histori.place(x=-1187, y=0, width=1187, height=812)
    impri.place(x=-1187, y=0, width=1187, height=812)
    produtosVendidos.place(x=-1187, y=0, width=1187, height=812)
    home.place(x=-1187, y=0, width=950, height=620)
    vendas.place(x=-1187, y=0, width=950, height=620)
    avancar.place(x=-1187, y=0, width=950, height=620)
    pesquisaproproambi.delete(0,END)
    nomepronovo.delete(0,END)
    qtdpronovo.delete(0,END)
    valoepronovo.delete(0,END)
    pesquisaambi.delete(0,END)
    novoambi.delete(0,END)
    pesquisaambi.focus()
    mostarproambiente()
    en = 30
    
app = Tk()
app.title('Gerenciador')
app.geometry('1187x750')
app.resizable(width=0,height=0)
app.iconbitmap(minhapasta+"\\ico.ico")

def enter(event):
    global en
    global vendapo
    if en == 1:
        if vendapo == 0:
            cod = codBarra.get()
            if cod != '':
                insertProduto()
            else :
                pass
        else:
            try:
                itens = produtosparavenda.selection()[0]
                valores = produtosparavenda.item(itens, 'values')
                insertProduto()
            except:
                pass
    elif en == 5:
        buscarHome()
    elif en == 3:
        buscarVend()
    elif en ==10:
        buscarimpri()
    elif en ==20:
        buscarHistorico()
    elif en ==30:
        buscarproambiente()
        am = novoambi.get()
        no = nomepronovo.get()
        qtd = qtdpronovo.get()
        va = valoepronovo.get()
        if no != '' and qtd != '' and va != '':
            cproambiente()


def qui():
    sys.exit(0)

barrademenu = Menu(app)

barrademenu.add_cascade(label='Home', command=frameHome )

barrademenu.add_cascade(label='Vender',command=frameVendas)

barrademenu.add_cascade(label='Itens por ambientes', command=frameAmbientes)

barrademenu.add_cascade(label='Lista imprimir', command=frameListaimprimir)

barrademenu.add_cascade(label='Produtos vendidos', command=frameProdutosVendidos)

barrademenu.add_cascade(label='Histórico', command=frameHistorico)

ns = ttk.Style()
ns.configure('Treeview', background='gray', foreground='black',
             rowheight=25, fieldbackground='white', borderwidth=1, relief='solid')
ns.map('Treeview', background=[('selected', '#0a0')],  
       foreground=[('selected', 'black')])

################ FRAME HOME #################

home = Frame(app, bg='#48d')
home.place(x=0, y=0, width=1187, height=812)
frase=Label(home, text='', bg='#48d', fg='#000',font=('arial', 18))
frase.place(x=550, y=19, width=312, height=38)
Label(home, text='Estoque de produtos:', bg='#48d', fg='#000',
      font=('arial', 18)).place(x=250, y=9, width=312, height=38)

produtos = ttk.Treeview(home, columns=(
    'qtd','Nome do produto', 'valorunidade','valortotal','cod','ambi'), show='headings')
produtos.column('qtd', minwidth=50,
                width=50, anchor=CENTER, stretch=NO)
produtos.column('Nome do produto', minwidth=325,
                width=325, anchor=CENTER, stretch=NO)
produtos.column('valorunidade', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtos.column('valortotal', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtos.column('cod', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)
produtos.column('ambi', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)

produtos.heading('qtd', text='QTD', anchor=CENTER)
produtos.heading('Nome do produto',text='Nome do produto', anchor=CENTER)
produtos.heading('valorunidade', text='Valor (R$)', anchor=CENTER)
produtos.heading('valortotal',text='Valor total (R$)', anchor=CENTER)
produtos.heading('cod',text='cod', anchor=CENTER)
produtos.heading('ambi',text='ambi', anchor=CENTER)

produtos.tag_configure('oddrow', background="lightblue")
produtos.tag_configure('evenrow', background="white")
produtos.tag_configure('red', background="orange")

scrollbarhome = Scrollbar(produtos)
scrollbarhome.pack(side=RIGHT, fill=Y)
produtos.config(yscrollcommand=scrollbarhome.set)
scrollbarhome.config(command=produtos.yview)

produtos.place(x=50, y=113, width=654, height=575)

Label(home, text='Nome do produto:', bg='#48d', fg='#000',
      font=('arial black', 12)).place(x=210, y=50, width=160, height=20)

pesquisapro = Entry(home, width=25, font=('arial', 13))
pesquisapro.place(x=165, y=75, width=270, height=31)

Label(home, text='Valor:', bg='#48d', fg='#000',
      font=('arial black', 12)).place(x=475, y=50, width=50, height=20)

pesquisapreco = Entry(home, width=25, font=('arial', 13))
pesquisapreco.place(x=450, y=75, width=100, height=31)

Label(home, text='QTD:', bg='#48d', fg='#000',
      font=('arial', 14)).place(x=750, y=375, width=63, height=38)

vezess = Entry(home, width=25, font=('arial', 13))
vezess.place(x=812, y=375, width=44, height=38)

vAuto3 = IntVar()

check = Checkbutton(home,text="Deletar permanente",variable=vAuto3,onvalue=1,offvalue=0, bd=1, relief='solid')
check.place(x=963, y=244, width=162, height=25)

def valortotProduto():
    vquery = 'SELECT SUM(valortotal) FROM produto'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        valor = i[0]
    if valor == None:
        valor = 0
    Label(home, text='Valor total no estoque:\nR$: '+str(round(valor, 2)), bg='#48d', fg='#000', font=(
        'arial', 18)).place(x=805, y=19, width=375, height=63)

def mostrarProHome():
    produtos.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtos.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtos.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtos.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtos.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    produtos.column('ambi', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)
    cor = 0
    produtos.delete(*produtos.get_children())
    vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod,ambi FROM produto'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[0] == 0:
                produtos.insert("", "end", values=i, tags=('red',))
            else:
                produtos.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[0] == 0:
                produtos.insert("", "end", values=i, tags=('red',))
            else:
                produtos.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    valortotProduto()
    pesquisapro.delete(0,END)
    pesquisapreco.delete(0,END)

def mostarOrdem():
    
    produtos.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtos.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtos.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtos.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtos.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)   
    produtos.column('ambi', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)           
    cor = 0
    produtos.delete(*produtos.get_children())
    vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod,ambi FROM produto ORDER BY nome'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[0] == 0:
                produtos.insert("", "end", values=i, tags=('red',))
            else:
                produtos.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[0] == 0:
                produtos.insert("", "end", values=i, tags=('red',))
            else:
                produtos.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
            
    valortotProduto()
    pesquisapro.delete(0,END)
    pesquisapreco.delete(0,END)

def buscarHome():
    produtos.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtos.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtos.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtos.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtos.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    produtos.column('ambi', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)
    cor = 0
    produtos.delete(*produtos.get_children())
    val = pesquisapreco.get()
    if val == "":
        pass
    else:
        val = val + ".0"
    vquery = "SELECT qtuni,nome,valoruni,valortotal,cod,ambi FROM produto WHERE nome LIKE '%" + \
        pesquisapro.get()+"%' AND valoruni LIKE '"+ val +"%'"
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[0] == 0:
                produtos.insert("", "end", values=i, tags=('red',))
            else:
                produtos.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[0] == 0:
                produtos.insert("", "end", values=i, tags=('red',))
            else:
                produtos.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    valortotProduto()
    
def alteraProduto():
    try:
        itens = produtos.selection()[0]
        valores = produtos.item(itens, 'values')
        try:
            telaAlteraProduto(valores[1],valores[4],valores[5])
        except:
            pass
    except:
        messagebox.showwarning(
            title='ERROR', message='Selecione o item da tabela.')

def deletarpro():
    try:
        itens = produtos.selection()[0]
        valores = produtos.item(itens, 'values')
        try:
            m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
            if m == True:
                Criar.deletaproduto(valores[4])
                Criar.deletaprodutoquasevv(valores[4])
                Criar.deletaprodutoquasevv2(valores[4])
                buscarHome()    
                valortotProduto()
        except:
            pass
    except:
        messagebox.showwarning(
            title='ERROR', message='Selecione o item da tabela.')

def recproTudo():
    threading.Thread(target=recproTudo2).start()

def recproTudo2():
    vquery = 'SELECT * FROM produtore'
    linhas = Criar.mostrar(vquery)
    linha = 0
    cont = 0
    for i in linhas:
        if cont == 0:
            frase.configure(text="Carregando.")
        elif cont == 2:
            frase.configure(text="Carregando..")
        elif cont == 3:
            frase.configure(text="Carregando...")
        elif cont == 4:
            cont = 0 
        if linha%2 == 0:
            cont +=1  
        Criar.cadastraprodo2(i[0],i[1],i[2],i[3],i[4],i[5],i[1],i[7])
        linha +=1

    Criar.deleTudoProrec()
    frase.configure(text="Aguarde")
    mostrarProHome()
    frase.configure(text="")

def deletarproTudo():
    threading.Thread(target=deletarproTudo2).start()

def deletarproTudo2():
    try:
        m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
        if m == True:
            a = vAuto3.get()
            if a == 0:
                linha = 0
                cont = 0
                vquery = 'SELECT * FROM produto'
                linhas = Criar.mostrar(vquery)
                for i in linhas:
                    if cont == 0:
                        frase.configure(text="Carregando.")
                    elif cont == 2:
                        frase.configure(text="Carregando..")
                    elif cont == 3:
                        frase.configure(text="Carregando...")
                    elif cont == 4:
                        cont = 0 
                    if linha%2 == 0:
                        cont +=1   
                    Criar.cadastraprodutore(i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7])  
                    linha +=1  
            Criar.deleTudoPro()
            Criar.deletaprodutotudoquase()
            Criar.deletaprodutotudoquase2()
            frase.configure(text="")
            mostrarProHome()
    except:
        pass

def deletarproTudozero():
    threading.Thread(target=deletarproTudozero2).start()

def deletarproTudozero2():
    try:
        m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
        if m == True:
            frase.configure(text='Aguarde')
            vquery = 'SELECT cod FROM produto WHERE qtuni = 0'
            linhas = Criar.mostrar(vquery)
            for i in linhas:
                Criar.deletaproduto(i[0])
            mostrarProHome()
            frase.configure(text='')
    except:
        pass

def abrirarq():
    global abr3
    if abr3 == 0:
        threading.Thread(target=abrirarq2).start()

def abrirarq2():
    try:
        di = filedialog.askopenfile()
        df = pd.read_excel(di.name)
        linha = 0
        cont = 0
        
    except:
        return
    try:
        numeros = len(df)
        for i in range(numeros):
            qtd = df['QTD'][linha]
            nome = df['Produto'][linha]
            valoruni = df['Valor uni'][linha]
            valortot = df['Valor total'][linha]
            cod = unidecode(nome) + " R$:"+str(valoruni)
            ambi = df['Ambiente'][linha]
            if cont == 0:
                frase.configure(text="Carregando.")
            elif cont == 2:
                frase.configure(text="Carregando..")
            elif cont == 3:
                frase.configure(text="Carregando...")
            elif cont == 4:
                cont = 0 
            if linha%2 == 0:
                cont +=1    
            nomes = str(nome.lower())
            Criar.cadastraproduto2(nomes,int(qtd),float(valoruni),float(valortot),cod,int(qtd),ambi.lower())
            linha +=1
        try:
            i1()
        except:
            messagebox.showerror(title='ERROR', message='Erro ao somar colunas iguais')
        try:
            i2()
        except:
            messagebox.showerror(title='ERROR', message='Erro ao adicionar na tabela principal')
        try:
            i3()
        except:
            messagebox.showerror(title='ERROR', message='Erro cadastrar ambientes')

    except:
        messagebox.showerror(title='ERROR', message='Insira as colunas na 1ª linha:\nQTD\nProduto\nValor uni\nValor total\nAmbiente\n\nRemova a soma total do final da planilha')
    
    Criar.deleTudoPro2()
    frase.configure(text="Aguarde")
    mostrarProHome()
    frase.configure(text="") 

def i1():
    linha = 0
    cont = 0
    vquery = 'SELECT nome,cod,Count(*) FROM produto2 GROUP BY cod HAVING Count(*) > 1' 
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        vquer = "SELECT nome,SUM(qtuni),valoruni,SUM(valortotal),cod,ambi FROM produto2 WHERE cod LIKE '"+i[1]+"'" 
        linha2 = Criar.mostrar(vquer)
        if cont == 0:
            frase.configure(text="Carregando.")
        elif cont == 2:
            frase.configure(text="Carregando..")
        elif cont == 3:
            frase.configure(text="Carregando...")
        elif cont == 4:
            cont = 0 
        if linha%2 == 0:
            cont +=2
        linha +=1
        Criar.deleproduto2(i[1]) 
        for j in linha2:
            Criar.cadastraproduto2(j[0],int(j[1]),float(j[2]),float(j[3]),j[4],int(j[1]),j[5]) 


def i2():
    linha = 0
    cont = 0
    vquery2 = 'SELECT * FROM produto2'
    linhas2 = Criar.mostrar(vquery2)
    for v in linhas2:
        codg = v[4]
        img = qrcode.make(codg)
        img.save('barcode.jpg')
        if cont == 0:
            frase.configure(text="Carregando.")
        elif cont == 2:
            frase.configure(text="Carregando..")
        elif cont == 3:
            frase.configure(text="Carregando...")
        elif cont == 4:
            cont = 0 
        if linha%2 == 0:
            cont +=1
        Criar.cadastraproduto(v[0],v[1],v[2],v[3],v[4],v[5],v[6])
        linha +=1 

def i3():
    vquery = 'SELECT DISTINCT ambi FROM produto ' 
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        Criar.cadastraambiente(i[0])
    jun()

################ FRAME VENDAS #################

vendas = Frame(app, bg='#48d')

c = Label(vendas, text='Cliente 1:', bg='#48d', fg='#000',font=('arial', 18))
c.place(x=125, y=38, width=125, height=38)

nomeCliente = Entry(vendas, width=25, font=('arial', 13))
nomeCliente.place(x=250, y=38, width=300, height=38)


Label(vendas, text='CEL:', bg='#48d', fg='#000',
      font=('arial', 18)).place(x=525, y=38, width=125, height=38)

celCliente = Entry(vendas, width=25, font=('arial', 13))
celCliente.place(x=625, y=38, width=300, height=38)

nomeCliente2 = Entry(vendas, width=25, font=('arial', 13))
celCliente2 = Entry(vendas, width=25, font=('arial', 13))


Label(vendas, text='Lista de produtos', bg='#48d', fg='#000',
      font=('arial', 18)).place(x=588, y=125, width=250, height=38)

labelcod = Label(vendas, text='', bg='#48d', fg='#000',
      font=('arial', 18))
labelcod.place(x=113, y=108, width=313, height=38)

codBarra = Entry(vendas, width=25, font=('arial', 13))

labelqtd = Label(vendas, text='QTD:', bg='#48d', fg='#000',
      font=('arial', 18))

qtdadd = Entry(vendas, width=25, font=('arial', 13))

pesquisavenda = Entry(vendas, width=25, font=('arial', 13))

def vai(event):
    pesquisaProvendas()

pesquisavenda.bind(sequence="<Return>",func=vai)

produtoVenda = ttk.Treeview(vendas, columns=(
    'qtd','Nome do produto', 'valorunidade','valortotal','cod','codes'), show='headings')
produtoVenda.column('qtd', minwidth=50,
                width=50, anchor=CENTER, stretch=NO)
produtoVenda.column('Nome do produto', minwidth=325,
                width=325, anchor=CENTER, stretch=NO)
produtoVenda.column('valorunidade', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtoVenda.column('valortotal', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtoVenda.column('cod', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)
produtoVenda.column('codes', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)

produtoVenda.heading('qtd', text='QTD', anchor=CENTER)
produtoVenda.heading('Nome do produto',text='Nome do produto', anchor=CENTER)
produtoVenda.heading('valorunidade', text='Valor (R$)', anchor=CENTER)
produtoVenda.heading('valortotal',text='Valor total (R$)', anchor=CENTER)
produtoVenda.heading('cod',text='cod', anchor=CENTER)
produtoVenda.heading('codes',text='codes', anchor=CENTER)

produtoVenda.tag_configure('oddrow', background="lightblue")
produtoVenda.tag_configure('evenrow', background="white")

produtoVenda.place(x=500, y=163, width=654, height=475)


produtosparavenda = ttk.Treeview(vendas, columns=(
    'nome','valor','cod'), show='headings')
produtosparavenda.column('nome', minwidth=285,
                width=285, anchor=CENTER, stretch=NO)
produtosparavenda.column('valor', minwidth=100,
                width=100, anchor=CENTER, stretch=NO)
produtosparavenda.column('cod', minwidth=0,
                width=0, anchor=CENTER, stretch=NO)

produtosparavenda.heading('nome', text='Nome', anchor=CENTER)
produtosparavenda.heading('valor',text='Valor', anchor=CENTER)
produtosparavenda.heading('cod', text='cod', anchor=CENTER)

produtosparavenda.tag_configure('oddrow', background="lightblue")
produtosparavenda.tag_configure('evenrow', background="white")
produtosparavenda.tag_configure('red', background="orange")

scrollbarprodutosparavenda = Scrollbar(produtosparavenda)
scrollbarprodutosparavenda.pack(side=RIGHT, fill=Y)
produtosparavenda.config(yscrollcommand=scrollbarprodutosparavenda.set)
scrollbarprodutosparavenda.config(command=produtosparavenda.yview)



def mostrarProvendas():
    produtosparavenda.column('nome', minwidth=285,
                    width=285, anchor=CENTER, stretch=NO)
    produtosparavenda.column('valor', minwidth=100,
                    width=100, anchor=CENTER, stretch=NO)
    produtosparavenda.column('cod', minwidth=0,
                    width=0, anchor=CENTER, stretch=NO)
    cor = 0
    produtosparavenda.delete(*produtosparavenda.get_children())
    vquery = 'SELECT nome,valoruni,cod FROM produto'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            produtosparavenda.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            produtosparavenda.insert("", "end", values=i, tags=('evenrow',))
            cor = 0

    pesquisavenda.delete(0,END)

def pesquisaProvendas():
    produtosparavenda.column('nome', minwidth=285,
                    width=285, anchor=CENTER, stretch=NO)
    produtosparavenda.column('valor', minwidth=100,
                    width=100, anchor=CENTER, stretch=NO)
    produtosparavenda.column('cod', minwidth=0,
                    width=0, anchor=CENTER, stretch=NO)
    cor = 0
    produtosparavenda.delete(*produtosparavenda.get_children())
    vquery = 'SELECT nome,valoruni,cod FROM produto WHERE nome LIKE "%'+pesquisavenda.get()+'%"'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            produtosparavenda.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            produtosparavenda.insert("", "end", values=i, tags=('evenrow',))
            cor = 0

def insertProduto():
    global cli
    global vendapo
    quant = qtdadd.get()
    if vendapo == 0:
        cod = codBarra.get()
    else:
        try:
            itens = produtosparavenda.selection()[0]
            valores = produtosparavenda.item(itens, 'values')
            cod = valores[2]
        except:
            messagebox.showerror(title="ERROR",message="Selecione um item da tabela")
            return
    qtd = 0
    nome = ''
    valorUni = 0    
    if quant == '':
        quant = 1
    if cod == '':
        messagebox.showwarning(title='ERROR', message="Preencha os dados")
    else:
        vquery = "SELECT caracteres,nome,valoruni,cod FROM produto WHERE cod LIKE '"+cod+"'"
        linhas = Criar.mostrar(vquery)
        if len(linhas) == 0:
            messagebox.showwarning(title='ERROR', message="Produto não encontrado")
            qtdadd.delete(0,END)
            codBarra.delete(0,END)
            codBarra.focus()
        else:
            for i in linhas:
                qtd = i[0]
                nome = i[1]
                valorUni = i[2]
                codigo = i[3]
            
            num = int(qtd) - int(quant)
            if num < 0:
                messagebox.showwarning(title='ERROR', message="Quantidade insuficiente no estoque\nEstoque:"+str(qtd))
                qtdadd.delete(0,END)
                codBarra.delete(0,END)
                codBarra.focus()
            else:
                valorTo = int(quant) * float(valorUni)
                cods = str(cod) + str(random())
                if cli == 1:
                    Criar.cadastraprodutoquase(nome,quant,valorUni,round(float(valorTo),2),cods,codigo)

                else:
                    Criar.cadastraprodutoquase2(nome,quant,valorUni,round(float(valorTo),2),cods,codigo)
                
                Criar.alteracarctere(int(num),str(codigo))
                  
                mostrarquaseVend()
                qtdadd.delete(0,END)
                codBarra.delete(0,END)
                codBarra.focus()

def incli1():
    global cli 
    cli = 1
    nomeCliente.place(x=250, y=38, width=300, height=38)
    celCliente.place(x=625, y=38, width=300, height=38)
    nomeCliente2.place(x=250, y=-100, width=313, height=38)
    celCliente2.place(x=625, y=-100, width=300, height=38)
    nomeCliente.focus()
    c.configure(text="Cliente 1:")
    mostrarquaseVend()

def incli2():
    global cli 
    cli = 2
    nomeCliente2.place(x=250, y=38, width=275, height=38)
    celCliente2.place(x=625, y=38, width=300, height=38)
    nomeCliente.place(x=250, y=-100, width=300, height=38)
    celCliente.place(x=625, y=-100, width=300, height=38)
    nomeCliente2.focus()
    c.configure(text="Cliente 2:")
    mostrarquaseVend()

def mostrarquaseVend():
    global cli
    produtoVenda.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtoVenda.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtoVenda.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoVenda.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoVenda.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    produtoVenda.column('codes', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)              
    cor = 0
    produtoVenda.delete(*produtoVenda.get_children())
    if cli == 1:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod,codes FROM produtosquasevendidos'
    else:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod,codes FROM produtosquasevendidos2'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            produtoVenda.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            produtoVenda.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    somaPP2()



def deleteQuase():
    global cli
    try:
        itens = produtoVenda.selection()[0]
        valores = produtoVenda.item(itens, 'values')
        try:
            m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
            if m == True:
                vquery = "SELECT caracteres FROM produto WHERE cod LIKE '"+valores[5]+"'"
                linhas = Criar.mostrar(vquery)
                for i in linhas:
                    num = int(i[0])+int(valores[0])
                    Criar.alteracarctere(num,valores[5])

                if cli == 1:
                    Criar.deletaprodutoquasevend(valores[4])
                else:
                    Criar.deletaprodutoquasevend2(valores[4])
                    
                mostrarquaseVend()
        except:
            pass
    except:
        messagebox.showwarning(
            title='ERROR', message='Selecione o item da tabela.')

def deletAlll():
    global cli
    if cli ==1:
        threading.Thread(target=deletAlll1).start()
    else:
        threading.Thread(target=deletAlll2).start()

def deletAlll1():
    global en
    if en == 1:
        m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
    else:
        m = True
    if m == True:
        vquery = 'SELECT cod,codes,Count(*) FROM produtosquasevendidos GROUP BY codes HAVING Count(*) > 1' 
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            vquer = "SELECT nome,SUM(qtuni),valoruni,SUM(valortotal),codes FROM produtosquasevendidos WHERE codes LIKE '"+i[1]+"'" 
            linha = Criar.mostrar(vquer)
            Criar.deletaprodutoquasevend20(i[1]) 
            for j in linha:
                Criar.cadastraprodutoquase(j[0],int(j[1]),j[2],float(j[3]),j[4]+str(random()),j[4])

        vquery20 = "SELECT produtosquasevendidos.codes, produtosquasevendidos.qtuni,produto.caracteres  \
            FROM produto, produtosquasevendidos WHERE produtosquasevendidos.codes = produto.cod "
        linhas20 = Criar.mostrar(vquery20)

        for ia in linhas20:
            soma = int(ia[2])+int(ia[1])
            Criar.alteracarctere(soma,ia[0])
        Criar.deletaprodutotudoquase()
        mostrarquaseVend()
    else:
        pass

def deletAlll2():
    global en
    if en == 1:
        m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
    else:
        m = True
    if m == True:
        vquery = 'SELECT cod,codes,Count(*) FROM produtosquasevendidos2 GROUP BY codes HAVING Count(*) > 1' 
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            vquer = "SELECT nome,SUM(qtuni),valoruni,SUM(valortotal),codes FROM produtosquasevendidos2 WHERE codes LIKE '"+i[1]+"'" 
            linha = Criar.mostrar(vquer)
            Criar.deletaprodutoquasevend40(i[1]) 
            for j in linha:
                Criar.cadastraprodutoquase2(j[0],int(j[1]),j[2],float(j[3]),j[4]+str(random()),j[4])


        vquery20 = "SELECT produtosquasevendidos2.codes, produtosquasevendidos2.qtuni,produto.caracteres \
            FROM produto JOIN produtosquasevendidos2 ON produtosquasevendidos2.codes = produto.cod "
        linhas20 = Criar.mostrar(vquery20)
        for ia in linhas20:
            soma = int(ia[2])+int(ia[1])
            Criar.alteracarctere(soma,ia[0])
        Criar.deletaprodutotudoquase2()
        mostrarquaseVend()
    else:
        pass


def somaPP2():
    global cli
    if cli == 1:
        vquery = 'SELECT SUM(valortotal) FROM produtosquasevendidos'
    else:
        vquery = 'SELECT SUM(valortotal) FROM produtosquasevendidos2'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        valor = i[0]
    if valor == None:
        valor = 0
    Label(vendas, text='Valor total a pagar:\nR$: '+str(round(valor, 2)), bg='#48d', fg='#000', font=(
        'arial', 18)).place(x=812, y=100, width=313, height=63)

################ FRAME AVANCAR #################

avancar = Frame(app, bg='#48d')

me = Label(avancar, text='', bg='#48d', fg='red', font=('arial', 18))
me.place(x=813, y=250, width=313, height=63)

v = Label(avancar, text='', bg='#48d', fg='#000', font=('arial', 18))
v.place(x=812, y=400, width=313, height=63)

valortotalsempre = Label(avancar, text='', bg='#48d', fg='#000', font=('arial', 18))
valortotalsempre.place(x=812, y=19, width=313, height=63)

Label(avancar, text='Desconto:', bg='#48d', fg='#000', font=(
    'arial', 18)).place(x=815, y=220, width=120, height=38)

des = Entry(avancar, width=25, font=('arial', 13))
des.place(x=935, y=220, width=65, height=38)


Label(avancar, text='Método de pagamento:', bg='#48d', fg='#000', font=(
    'arial', 18)).place(x=790, y=320, width=250, height=38)

metodosdepag=["Selecione","Débito","Crédito","Pix","Transferência","Dinheiro"]

vmetodo = StringVar()

metodo = OptionMenu(avancar,vmetodo,metodosdepag[1],metodosdepag[2],metodosdepag[3],metodosdepag[4],metodosdepag[5])
metodo.place(x=1045, y=320, width=120, height=38)

Label(avancar, text='Status:', bg='#48d', fg='#000', font=(
    'arial', 18)).place(x=790, y=480, width=250, height=38)

statuspag=["Selecione","Pago","Pendente"]

vstatus = StringVar()

status = OptionMenu(avancar,vstatus,statuspag[1],statuspag[2])
status.place(x=980, y=480, width=120, height=38)

vAuto2 = IntVar()

check2 = Checkbutton(avancar,text="Abrir whatsapp",variable=vAuto2,onvalue=1,offvalue=0, bd=1, relief='solid')



def som():
    global cli
    Label(avancar, text='Carregando...',bg="#48d", fg='red',font=('arial', 18)).place(x=200, y=250, width=250, height=30)
    if cli ==1:
        vquery = 'SELECT cod,codes,Count(*) FROM produtosquasevendidos GROUP BY codes HAVING Count(*) > 1' 
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            vquer = "SELECT nome,SUM(qtuni),valoruni,SUM(valortotal),codes FROM produtosquasevendidos WHERE codes LIKE '"+i[1]+"'" 
            linha = Criar.mostrar(vquer)
            Criar.deletaprodutoquasevend20(i[1]) 
            for j in linha:
                Criar.cadastraprodutoquase(j[0],int(j[1]),j[2],float(j[3]),j[4]+str(random()),j[4])
    else:
        vquery = 'SELECT cod,codes,Count(*) FROM produtosquasevendidos2 GROUP BY codes HAVING Count(*) > 1' 
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            vquer = "SELECT nome,SUM(qtuni),valoruni,SUM(valortotal),codes FROM produtosquasevendidos2 WHERE codes LIKE '"+i[1]+"'" 
            linha = Criar.mostrar(vquer)
            Criar.deletaprodutoquasevend40(i[1]) 
            for j in linha:
                Criar.cadastraprodutoquase2(j[0],int(j[1]),j[2],float(j[3]),j[4]+str(random()),j[4])
    m()


def m():
    global cli
    global nomedocli
    global numerodocle
    a = Frame(avancar, bg='white')
    scrollbar = Scrollbar(a)
    scrollbar.pack(side=RIGHT, fill=Y)
    textbox = Text(a)
    if cli ==1:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod FROM produtosquasevendidos'
    else:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod FROM produtosquasevendidos2'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        textbox.insert(END,f"  Produto:{i[1]}\n  Quantidade:{i[0]}\n  Valor unitario:{i[2]}\n  Valor total:{i[3]}\n#####################################\n")
    a.place(x=125, y=125, width=654, height=563)
    textbox.pack()
    textbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=textbox.yview)
    textbox.configure(state='disabled',font=('Arial',14))
    Label(avancar, text='Cliente: '+nomedocli, bg='#48d', fg='#000', font=(
        'arial', 14),anchor=W).place(x=125, y=19, width=375, height=38)
    celular = numerodocle
    if len(celular) == 10:
        telFormatado = '({}) {}{}-{}'.format(celular[0:2],celular[2] ,celular[3:6], celular[6:]) 
    else:
        telFormatado = '({}) {}{}-{}'.format(celular[0:2],celular[2] ,celular[3:7], celular[7:]) 
    Label(avancar, text='Cel: '+ telFormatado, bg='#48d', fg='#000', font=(
        'arial', 14),anchor=W).place(x=125, y=63, width=375, height=38)
    somaPP()

def remodesconto():
    global ativades
    v.configure(text="")
    somaPP()
    ativades = 0

def calva():
    global ativades
    global valo
    somaPP()
    va = des.get()
    if va == '':
        return
    try:
        valo = valo - int(va)
        ativades = 1
        v.configure(text='Valor com desconto:\nR$: '+str(valo))
    except:
        messagebox.showwarning(title='ERROR', message='Insira um valor numerico')


def calpor():
    global ativades
    global quantdodes
    global valo
    somaPP()
    va = des.get()
    if va == '':
        return
    try:
        por = int(va)* valo / 100
        valo = valo - por
        ativades = 2
        quantdodes = por
        v.configure(text='Valor com desconto:\nR$: '+str(valo))
    except:
        messagebox.showwarning(title='ERROR', message='Insira um valor numerico')

def somaPP():
    global cli
    global valo
    global sempre
    global ativades
    if cli ==1:
        vquery = 'SELECT SUM(valortotal) FROM produtosquasevendidos'
    else:
        vquery = 'SELECT SUM(valortotal) FROM produtosquasevendidos2'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        valor = i[0]
    if valor == None:
        valor = 0
    valo = round(valor, 2)
    sempre = valo
    ativades = 0
    valortotalsempre.configure(text='Valor total a pagar:\nR$: '+str(valo))


def criarRegistro2():
    global cli
    global nomedocli
    global numerodocle
    global valo
    global ativades
    global quantdodes
    global sempre
    if os.path.isdir(minhapasta + '\\Enviar'):
        dir = os.listdir(minhapasta + '\\Enviar')
        for file in dir:
            os.remove(minhapasta + '\\Enviar\\'+file)
    else:
        os.mkdir(minhapasta + '\\Enviar')
    if cli ==1:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal FROM produtosquasevendidos'
    else:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal FROM produtosquasevendidos2'
    linhas = Criar.mostrar(vquery)
    arquivo = xlsxwriter.Workbook(minhapasta+"\\Enviar\\"+nomedocli+".xlsx")
    table = arquivo.add_worksheet()
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format.set_align('vcenter')
    currency_format.set_font_size(32)
    currency_format.set_num_format('R$ #,##0.00')
    metododepag = vmetodo.get()
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format2.set_align('vcenter')
    currency_format2.set_font_size(42)

    currency_format3 = arquivo.add_format()
    currency_format3.set_align('center')
    currency_format3.set_align('vcenter')
    currency_format3.set_font_size(32)

    currency_format4 = arquivo.add_format()
    currency_format4.set_align('center')
    currency_format4.set_align('vcenter')
    currency_format4.set_font_size(30)
    currency_format4.set_text_wrap()

    table.set_column_pixels("A:A",width=100)
    table.set_column_pixels("B:B",width=420)
    table.set_column_pixels("C:C",width=290)
    table.set_column_pixels("D:D",width=310)
    table.set_row_pixels(0,50)

    data = datetime.now().strftime('%d/%m/%Y')
    table.write(0,0,"Data:",currency_format3)
    table.write(0,1,data,currency_format2)

    table.write(0,2,"Pagamento:",currency_format3)
    table.write(0,3,metododepag,currency_format2)


    table.set_row_pixels(1,50)
    table.set_row_pixels(2,50)
    table.set_row_pixels(3,50)
    table.write(2,0,"QTD",currency_format3)
    table.write(2,1,"Produto",currency_format2)
    table.write(2,2,"Valor unidade",currency_format3)
    table.write(2,3,"Valor total",currency_format3)
    
    row = 4
    col = 0
    for i in linhas:
        table.set_row_pixels(row,100)
        table.write(row,col, i[0],currency_format2)
        table.write(row,col +1, i[1],currency_format4)
        table.write(row,col +2,i[2],currency_format)
        table.write(row,col +3, i[3],currency_format)
        row +=1

    table.write(row+1,col +2,"Valor total:",currency_format3)
    table.write(row+1,col +3,sempre,currency_format)

    if ativades == 0:
        pass
    elif ativades == 1:
        table.write(row+2,col +2,"Desconto:",currency_format3)
        table.write(row+2,col +3,float(des.get())*-1,currency_format)

        table.write(row+3,col +2,"Valor final:",currency_format3)
        table.write(row+3,col +3,valo,currency_format)

    elif ativades == 2:
        table.write(row+2,col +2,"Desconto:",currency_format3)
        table.write(row+2,col +3,float(quantdodes)*-1,currency_format)

        table.write(row+3,col +2,"Valor final:",currency_format3)
        table.write(row+3,col +3,valo,currency_format)

    arquivo.close()

def criarRegistro():
    global cli
    global nomedocli
    global numerodocle
    global valo
    global ativades
    global quantdodes
    global sempre
    nomearquivo = nomedocli+"_"+str(int(valo)) + str(round(float(random()),7))
    if os.path.isdir(minhapasta + '\\Meu estoque\\Vendas'):
        pass
    else:
        os.mkdir(minhapasta + '\\Meu estoque\\Vendas')
    try:
        os.remove(minhapasta+"\\Meu estoque\\Vendas\\"+nomearquivo+".xlsx")
    except:
        pass
    if cli ==1:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal FROM produtosquasevendidos'
    else:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal FROM produtosquasevendidos2'
    linhas = Criar.mostrar(vquery)
    arquivo = xlsxwriter.Workbook(minhapasta+"\\Meu estoque\\Vendas\\"+nomearquivo+".xlsx")
    table = arquivo.add_worksheet()
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format.set_align('vcenter')
    currency_format.set_num_format('R$ #,##0.00')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format2.set_align('vcenter')
    currency_format2.set_align('top')
    currency_format3 = arquivo.add_format()
    currency_format3.set_align('right')
    currency_format3.set_align('vcenter')
    table.set_column_pixels("A:A",width=50)
    table.set_column_pixels("B:B",width=266)
    table.set_column_pixels("C:C",width=120)
    table.set_column_pixels("D:D",width=150)
    table.write(0,0,"Cliente:",currency_format2)
    table.write(0,1,nomedocli)
    table.write(0,2,"Cel:",currency_format3)
    celular = numerodocle
    if len(celular) == 10:
        telFormatado = '({}) {}{}-{}'.format(celular[0:2],celular[2] ,celular[3:6], celular[6:]) 
    else:
        telFormatado = '({}) {}{}-{}'.format(celular[0:2],celular[2] ,celular[3:7], celular[7:])  
    
    metododepag = vmetodo.get()
    table.write(0,3,telFormatado,currency_format2)
    data = datetime.now().strftime('%d/%m/%Y')
    table.write(1,2,"Data:",currency_format3)
    table.write(1,3,data,currency_format2)

    table.write(2,2,"Pagamento:",currency_format3)
    table.write(2,3,metododepag,currency_format2)

    table.write(4,0,"QTD",currency_format2)
    table.write(4,1,"Produto",currency_format2)
    table.write(4,2,"Valor unidade",currency_format2)
    table.write(4,3,"Valor total",currency_format2)
    row = 6
    col = 0
    for i in linhas:
        table.write(row,col, i[0],currency_format2)
        table.write(row,col +1, i[1],currency_format2)
        table.write(row,col +2,i[2],currency_format)
        table.write(row,col +3, i[3],currency_format)
        row +=1


    table.write(row+1,col +2,"Valor total:",currency_format3)
    table.write(row+1,col +3,sempre,currency_format)

    if ativades == 0:
        pass
    elif ativades == 1:
        table.write(row+2,col +2,"Desconto:",currency_format3)
        table.write(row+2,col +3,float(des.get())*-1,currency_format)

        table.write(row+3,col +2,"Valor final:",currency_format3)
        table.write(row+3,col +3,valo,currency_format)

    elif ativades == 2:
        table.write(row+2,col +2,"Desconto:",currency_format3)
        table.write(row+2,col +3,float(quantdodes)*-1,currency_format)

        table.write(row+3,col +2,"Valor final:",currency_format3)
        table.write(row+3,col +3,valo,currency_format)

    arquivo.close()
    st = vstatus.get()
    Criar.cadastrahist(nomedocli,data,metododepag,valo,st,nomearquivo)


def agrupar():
    vquery = 'SELECT cod,codes,Count(*) FROM produtosvendidos GROUP BY codes HAVING Count(*) > 1' 
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        vquer = "SELECT nome,SUM(qtuni),valoruni,SUM(valortotal),codes FROM produtosvendidos WHERE codes LIKE '"+i[1]+"'" 
        linha = Criar.mostrar(vquer)
        Criar.deletaprodutovend2(i[1]) 
        for j in linha:
            Criar.cadastraprodutoVendido(j[0],int(j[1]),j[2],float(j[3]),j[4]+str(random()),j[4])    

def addVendi():
    global cli
    global valo
    global ativades
    global nomedocli
    global quantdodes
    if cli ==1:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod,codes FROM produtosquasevendidos'
    else:
        vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod,codes FROM produtosquasevendidos2'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        Criar.cadastraprodutoVendido(i[1],i[0],i[2],i[3],i[4],i[5])
        valo = 0

    if ativades == 1:
        va = des.get()
        desconto = float(va)*-1
        Criar.cadastraprodutoVendido("Desconto "+str(nomedocli),1,desconto,desconto,"Desconto "+str(nomedocli)+str(desconto) + str(random()),"Desconto "+str(nomedocli)+str(desconto))
    elif ativades == 2:
        desconto = float(quantdodes)*-1
        Criar.cadastraprodutoVendido("Desconto "+str(nomedocli),1,desconto,desconto,"Desconto "+str(nomedocli)+str(desconto) + str(random()),"Desconto "+str(nomedocli)+str(desconto))
    else:
        pass

def retiraVendi():
    global cli
    if cli ==1:
        vquery = "SELECT produtosquasevendidos.codes,produtosquasevendidos.qtuni, produto.qtuni, produto.valoruni \
            FROM produto JOIN produtosquasevendidos ON produtosquasevendidos.codes = produto.cod "
    else:
        vquery = "SELECT produtosquasevendidos2.codes,produtosquasevendidos2.qtuni, produto.qtuni, produto.valoruni \
            FROM produto JOIN produtosquasevendidos2 ON produtosquasevendidos2.codes = produto.cod "
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        quantidade = int(i[2]) - int(i[1])
        valort = quantidade * float(i[3])
        Criar.vender(int(quantidade),round(float(valort),2),i[0])

    if cli ==1:
        Criar.deletaprodutotudoquase()
    else:
        Criar.deletaprodutotudoquase2()

def finaliza2():
    tip = vmetodo.get()
    st = vstatus.get()
    if tip == "Selecione":
        messagebox.showwarning(title='ERROR', message='Selecione o metodo de pagamento')
    elif st == "Selecione":
        messagebox.showwarning(title='ERROR', message='Selecione o status da venda')
    else:
        me.configure(text='Aguarde...')
        threading.Thread(target=finaliza).start()
    
def reg2():
    try:
        import pywhatkit
        pywhatkit.sendwhatmsg_instantly("+55"+numerodocle,"")
        os.startfile(minhapasta+"\\Enviar")
    except:
        messagebox.showwarning(title='ERROR', message='Sem internet')
        me.configure(text='')
        

def finaliza():
    global cli
    global numerodocle
    a= vAuto2.get()
    '''if a == 1:
        criarRegistro2()
        threading.Thread(target=reg2).start()'''
    criarRegistro()
    addVendi()
    agrupar()
    retiraVendi()
    frameProdutosVendidos()
    me.configure(text='')
    if cli ==1:
        nomeCliente.delete(0,END)
        celCliente.delete(0,END)
    else:
        nomeCliente2.delete(0,END)
        celCliente2.delete(0,END)
        
################ FRAME PRODUTOS VENDIDOS #################


produtosVendidos = Frame(app, bg='#48d')



Label(produtosVendidos, text='Produtos vendidos:', bg='#48d', fg='#000',
      font=('arial', 18)).place(x=250, y=19, width=313, height=38)

produtoVendido = ttk.Treeview(produtosVendidos, columns=(
    'qtd','Nome do produto', 'valorunidade','valortotal','cod'), show='headings')
produtoVendido.column('qtd', minwidth=50,
                width=50, anchor=CENTER, stretch=NO)
produtoVendido.column('Nome do produto', minwidth=325,
                width=325, anchor=CENTER, stretch=NO)
produtoVendido.column('valorunidade', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtoVendido.column('valortotal', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtoVendido.column('cod', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)
                
produtoVendido.heading('qtd', text='QTD', anchor=CENTER)
produtoVendido.heading('Nome do produto',text='Nome do produto', anchor=CENTER)
produtoVendido.heading('valorunidade', text='Valor (R$)', anchor=CENTER)
produtoVendido.heading('valortotal',text='Valor total (R$)', anchor=CENTER)
produtoVendido.heading('cod',text='cod', anchor=CENTER)

produtoVendido.tag_configure('oddrow', background="lightblue")
produtoVendido.tag_configure('evenrow', background="white")

scrollbarvendidos = Scrollbar(produtoVendido)
scrollbarvendidos.pack(side=RIGHT, fill=Y)
produtoVendido.config(yscrollcommand=scrollbarvendidos.set)
scrollbarvendidos.config(command=produtoVendido.yview)

produtoVendido.place(x=50, y=113, width=654, height=575)

pesquisaproVendido = Entry(produtosVendidos, width=25, font=('arial', 13))
pesquisaproVendido.place(x=250, y=75, width=300, height=31)


def buscarVend():
    produtoVendido.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtoVendido.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtoVendido.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoVendido.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoVendido.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)             
    cor = 0
    produtoVendido.delete(*produtoVendido.get_children())
    vquery = "SELECT qtuni,nome,valoruni,valortotal,cod FROM produtosvendidos WHERE nome LIKE '%" + \
        pesquisaproVendido.get()+"%'"
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            produtoVendido.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            produtoVendido.insert("", "end", values=i, tags=('evenrow',))
            cor = 0

def mostrarVend():
    produtoVendido.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtoVendido.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtoVendido.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoVendido.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoVendido.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    cor = 0
    produtoVendido.delete(*produtoVendido.get_children())
    vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod FROM produtosvendidos'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            produtoVendido.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            produtoVendido.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    valortotProdutoVend()

def valortotProdutoVend():
    vquery = 'SELECT SUM(valortotal) FROM produtosvendidos'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        valor = i[0]
    if valor == None:
        valor = 0
    Label(produtosVendidos, text='Valor total vendido:\nR$: '+str(round(valor, 2)), bg='#48d', fg='#000', font=(
        'arial', 18)).place(x=775, y=19, width=375, height=63)

def mostarOrdemVend():
    produtoVendido.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtoVendido.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtoVendido.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoVendido.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoVendido.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    cor = 0
    produtoVendido.delete(*produtoVendido.get_children())
    vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod FROM produtosvendidos ORDER BY nome'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            produtoVendido.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            produtoVendido.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    valortotProdutoVend()

def criarexcelbarrasVendidos2OA():
    global abr7
    if abr7 == 0:
        threading.Thread(target=criarexcelbarrasVendidos2OA2).start()
        abr7 = 1

def criarexcelbarrasVendidos2OA2():
    global abr7
    try:
        os.remove(minhapasta+ '\\Meu estoque\\Produtos_vendidos_alfabetica.xlsx')
    except:
        pass
    vquery = 'SELECT qtuni,nome,valoruni,valortotal FROM produtosvendidos ORDER BY nome'
    linhas = Criar.mostrar(vquery)
    arquivo = xlsxwriter.Workbook(minhapasta+ '\\Meu estoque\\Produtos_vendidos_alfabetica.xlsx')
    table = arquivo.add_worksheet()
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format.set_align('vcenter')
    currency_format.set_num_format('R$#,##0.00')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format2.set_align('vcenter')
    table.set_column_pixels("A:A",width=50)
    table.set_column_pixels("B:B",width=286)
    table.set_column_pixels("C:D",width=120)
    table.write(0,0,"QTD",currency_format2)
    table.write(0,1,"Produto",currency_format2)
    table.write(0,2,"Valor uni",currency_format2)
    table.write(0,3,"Valor total",currency_format2)
    row = 1
    col = 0
    for i in linhas:
        table.write(row,col, i[0],currency_format2)
        table.write(row,col +1, i[1],currency_format2)
        table.write(row,col +2,i[2],currency_format)
        table.write(row,col +3, i[3],currency_format)
        row +=1

    vquery2 = 'SELECT SUM(valortotal) FROM produtosvendidos'
    linhas2 = Criar.mostrar(vquery2)
    for i in linhas2:
        table.write(row,col +3,i[0],currency_format)

    arquivo.close()
    os.startfile(minhapasta+"\\Meu estoque\\Produtos_vendidos_alfabetica.xlsx")
    abr7 = 0


def criarexcelbarrasVendidos():
    global abr9
    if abr9 == 0:
        threading.Thread(target=criarexcelbarrasVendidos2).start()
        abr9 = 1

def criarexcelbarrasVendidos2():
    global abr9
    try:
        os.remove(minhapasta+ '\\Meu estoque\\Produtos_vendidos.xlsx')
    except:
        pass
    vquery = 'SELECT qtuni,nome,valoruni,valortotal FROM produtosvendidos'
    linhas = Criar.mostrar(vquery)
    arquivo = xlsxwriter.Workbook(minhapasta+ '\\Meu estoque\\Produtos_vendidos.xlsx')
    table = arquivo.add_worksheet()
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format.set_align('vcenter')
    currency_format.set_num_format('R$#,##0.00')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format2.set_align('vcenter')
    table.set_column_pixels("A:A",width=50)
    table.set_column_pixels("B:B",width=286)
    table.set_column_pixels("C:D",width=120)
    table.write(0,0,"QTD",currency_format2)
    table.write(0,1,"Produto",currency_format2)
    table.write(0,2,"Valor uni",currency_format2)
    table.write(0,3,"Valor total",currency_format2)
    row = 1
    col = 0
    for i in linhas:
        table.write(row,col, i[0],currency_format2)
        table.write(row,col +1, i[1],currency_format2)
        table.write(row,col +2,i[2],currency_format)
        table.write(row,col +3, i[3],currency_format)
        row +=1

    vquery2 = 'SELECT SUM(valortotal) FROM produtosvendidos'
    linhas2 = Criar.mostrar(vquery2)
    for i in linhas2:
        table.write(row,col +3,i[0],currency_format)

    arquivo.close()
    os.startfile(minhapasta+"\\Meu estoque\\Produtos_vendidos.xlsx")
    abr9 = 0

def prodelitem():
    try:
        itens = produtoVendido.selection()[0]
        valores = produtoVendido.item(itens, 'values')
        try:
            m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
            if m == True:
                Criar.deletaprodutovend(valores[4])
            mostrarVend()
        except:
            pass
    except:
        messagebox.askyesno(title='Aviso', message='Selecione um item')

def delintemVendidotudo():
    threading.Thread(target=delintemVendidotudo2).start()

def delintemVendidotudo2():
        try:
            m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
            if m == True:
                Criar.deltud()
                mostrarVend()
        except:
            pass


def delintemimpri():
    try:
        itens = produtosimpri.selection()[0]
        valores = produtosimpri.item(itens, 'values')
        try:
            m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
            if m == True:
                Criar.deletaprodutoimpri(valores[4])
                mostrarimpri()
        except:
            pass
    except:
        messagebox.showwarning(
            title='ERROR', message='Selecione o item da tabela.')

def delintemimpritudo():
    global abr3
    if abr3 ==0:
        threading.Thread(target=delintemimpritudo2).start()
        abr3 = 1

def delintemimpritudo2():
    global abr3
    try:
        m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
        if m == True:
            Criar.deltudimpri()
            mostrarimpri()
            abr3 = 0
    except:
        abr3 = 0
        pass

################ Frame lista imprimir #################

impri = Frame(app, bg='#48d')

Label(impri, text='Lista para imprimir:', bg='#48d', fg='#000',
      font=('arial', 18)).place(x=150, y=9, width=312, height=38)
fraseim=Label(impri, text='', bg='#48d', fg='#000',font=('arial', 18))
fraseim.place(x=430, y=19, width=312, height=38)

pagimpri=Label(impri, text='Pág completas para imprimir:', bg='#48d', fg='#000',font=('arial', 16))
pagimpri.place(x=720, y=25, width=280, height=42)

quantimpri=Label(impri, text='QTD de itens:', bg='#48d', fg='#000',font=('arial', 16))
quantimpri.place(x=1000, y=25, width=200, height=42)

conimpri=Label(impri, text='Itens para completar a pág:', bg='#48d', fg='#000',font=('arial', 16))
conimpri.place(x=820, y=100, width=310, height=42)

produtosimpri = ttk.Treeview(impri, columns=(
    'qtd','Nome do produto', 'valorunidade','valortotal','cod'), show='headings')
produtosimpri.column('qtd', minwidth=50,
                width=50, anchor=CENTER, stretch=NO)
produtosimpri.column('Nome do produto', minwidth=325,
                width=325, anchor=CENTER, stretch=NO)
produtosimpri.column('valorunidade', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtosimpri.column('valortotal', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtosimpri.column('cod', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)

produtosimpri.heading('qtd', text='QTD', anchor=CENTER)
produtosimpri.heading('Nome do produto',text='Nome do produto', anchor=CENTER)
produtosimpri.heading('valorunidade', text='Valor (R$)', anchor=CENTER)
produtosimpri.heading('valortotal',text='Valor total (R$)', anchor=CENTER)
produtosimpri.heading('cod',text='cod', anchor=CENTER)

produtosimpri.tag_configure('oddrow', background="lightblue")
produtosimpri.tag_configure('evenrow', background="white")
produtosimpri.tag_configure('red', background="orange")

scrollbarhome = Scrollbar(produtosimpri)
scrollbarhome.pack(side=RIGHT, fill=Y)
produtosimpri.config(yscrollcommand=scrollbarhome.set)
scrollbarhome.config(command=produtosimpri.yview)


produtosimpri.place(x=50, y=113, width=654, height=575)
  
pesquisaproImpri = Entry(impri, width=25, font=('arial', 13))
pesquisaproImpri.place(x=250, y=75, width=300, height=31)


def buscarimpri():
    produtosimpri.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtosimpri.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtosimpri.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtosimpri.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtosimpri.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)             
    cor = 0
    produtosimpri.delete(*produtosimpri.get_children())
    vquery = "SELECT qtuni,nome,valoruni,valortotal,cod FROM imprimir WHERE nome LIKE '%" + \
        pesquisaproImpri.get()+"%'"
    linhas = Criar.mostrar(vquery)
    cont = 1
    for i in linhas:
        if cont == 18:
            produtosimpri.insert("", "end", values=i, tags=('red',))
            cont =0
        elif cor == 0:
            produtosimpri.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            produtosimpri.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
        cont +=1
        

def mostrarimpri():
    produtosimpri.column('qtd', minwidth=50,
                    width=50, anchor=CENTER, stretch=NO)
    produtosimpri.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtosimpri.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtosimpri.column('valortotal', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtosimpri.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    cor = 0
    produtosimpri.delete(*produtosimpri.get_children())
    vquery = 'SELECT qtuni,nome,valoruni,valortotal,cod FROM imprimir'
    linhas = Criar.mostrar(vquery)
    cont = 1
    num = 0
    iten= 0
    b = 1
    pag = 0
    for i in linhas:
        if cont == 27:
            produtosimpri.insert("", "end", values=i, tags=('red',))
            b = 0
            cont =0
        elif cor == 0:
            produtosimpri.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            produtosimpri.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
        if cont ==0:
            pag+=1
        cont +=1
        num +=1
        iten = 27 - b
        b +=1
        
    if iten == 0:
        iten =27
    quantimpri.configure(text='QTD de itens:\n'+str(num))
    pagimpri.configure(text='Pág completas para imprimir:\n'+str(pag))
    conimpri.configure(text='Itens para completar a pág:\n'+str(iten))

def criarCoddebarras2impri():
    global abr3
    if abr3 ==0:
        threading.Thread(target=criarCoddebarrasimpri).start()
        abr3 = 1

def criarCoddebarrasimpri():
    global abr3
    try:
        if os.path.isdir(minhapasta+"\\img"):
            dir = os.listdir(minhapasta + '\\img')
            for file in dir:
                os.remove(minhapasta + '\\img\\'+file)
        else:
            os.mkdir(minhapasta+"\\img")
    except:
        pass
    try:
        os.remove(minhapasta+"\\Meu estoque\\QRCode.xlsx")
    except:
        pass
    vquery = 'SELECT nome,codimg,valoruni,cod FROM imprimir'
    linhas = Criar.mostrar(vquery)
    try:
        arquivo = xlsxwriter.Workbook(minhapasta+"\\Meu estoque\\QRCode.xlsx")
    except:
        return
    table = arquivo.add_worksheet()
    table.set_column_pixels("A:C",width=240)
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format3 = arquivo.add_format()
    currency_format3.set_align('center')
    currency_format4 = arquivo.add_format()
    currency_format4.set_align('center')
    currency_format.set_font_size(12)
    currency_format2.set_font_size(10)
    currency_format3.set_font_size(8)
    currency_format4.set_font_size(7)
    row = 1
    row2 = 2
    row3 = 0
    col = 0
    nun = 1
    cont = 0
    for i in linhas:
        img = i[1]
        with open(minhapasta+"\\img\\"+str(nun)+".jpg","wb") as f:
            f.write(img)
        ver = Image.open(minhapasta+"\\img\\"+str(nun)+".jpg")
        sa =ver.resize([200,150],Image.ANTIALIAS)
        sa.save(minhapasta+"\\img\\"+str(nun)+".jpg",quality=100)
        table.set_row_pixels(row,150)
        table.insert_image(row,col,minhapasta+"\\img\\"+str(nun)+".jpg",{'x_offset': 20})
        table.set_row_pixels(row2,16)

        table.set_row_pixels(row3,15)


        if int(len(str(i[3]))) <= 24:
            table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format)
        elif int(len(str(i[3]))) <= 28:
            table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format2)
        elif int(len(str(i[3]))) <= 45:
            table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format3)
        elif int(len(str(i[3]))) > 45:
            table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format4)
        if cont == 0:
            fraseim.configure(text="Carregando.")
        elif cont == 20:
            fraseim.configure(text="Carregando..")
        elif cont == 40:
            fraseim.configure(text="Carregando...")
        elif cont == 60:
            cont = 0   

        col +=1
        if col == 3:
            row +=3
            row2 +=3
            row3 +=3
            col =0
            cont +=1


        nun += 1

    fraseim.configure(text="Aguarde")   
    try:     
        arquivo.close() 
        os.startfile(minhapasta+"\\Meu estoque\\QRCode.xlsx")
    except:
        pass
    fraseim.configure(text="")
    remover()
    abr3 = 0


################ FRAME HISTORICO #################


histori = Frame(app, bg='#48d')

Label(histori, text='Histórico de vendas:', bg='#48d', fg='#000',
      font=('arial', 18)).place(x=150, y=9, width=312, height=38)

histpagov=Label(histori, text='Valor pago:\n R$:0', bg='#48d', fg='#000',font=('arial', 16))
histpagov.place(x=750, y=25, width=200, height=42)

histpendentev=Label(histori, text='Valor pendente:\n R$:0', bg='#48d', fg='#000',font=('arial', 16))
histpendentev.place(x=970, y=25, width=200, height=42)

listpessoas = ttk.Treeview(histori, columns=(
    'pessoa','data', 'forma','valortotal','status','cod'), show='headings')
listpessoas.column('pessoa', minwidth=288,
                width=288, anchor=CENTER, stretch=NO)
listpessoas.column('data', minwidth=100,
                width=100, anchor=CENTER, stretch=NO)
listpessoas.column('forma', minwidth=125,
                width=125, anchor=CENTER, stretch=NO)
listpessoas.column('valortotal', minwidth=100,
                width=100, anchor=CENTER, stretch=NO)
listpessoas.column('status', minwidth=125,
                width=125, anchor=CENTER, stretch=NO)
listpessoas.column('cod', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)

listpessoas.heading('pessoa', text='Cliente', anchor=CENTER)
listpessoas.heading('data',text='Data', anchor=CENTER)
listpessoas.heading('forma', text='Pagamento', anchor=CENTER)
listpessoas.heading('valortotal',text='Valor (R$)', anchor=CENTER)
listpessoas.heading('status',text='Status', anchor=CENTER)
listpessoas.heading('cod',text='cod', anchor=CENTER)

listpessoas.tag_configure('oddrow', background="lightblue")
listpessoas.tag_configure('evenrow', background="white")
listpessoas.tag_configure('red', background="orange")

scrollbarhome = Scrollbar(listpessoas)
scrollbarhome.pack(side=RIGHT, fill=Y)
listpessoas.config(yscrollcommand=scrollbarhome.set)
scrollbarhome.config(command=listpessoas.yview)


listpessoas.place(x=50, y=113, width=754, height=575)
  
pesquisaprohist = Entry(histori, width=25, font=('arial', 13))
pesquisaprohist.place(x=250, y=75, width=300, height=31)


def mostarOrdemhis():
    global ordenarpor
    listpessoas.column('pessoa', minwidth=288,
                    width=288, anchor=CENTER, stretch=NO)
    listpessoas.column('data', minwidth=100,
                    width=100, anchor=CENTER, stretch=NO)
    listpessoas.column('forma', minwidth=125,
                    width=125, anchor=CENTER, stretch=NO)
    listpessoas.column('valortotal', minwidth=100,
                    width=100, anchor=CENTER, stretch=NO)
    listpessoas.column('status', minwidth=125,
                    width=125, anchor=CENTER, stretch=NO)
    listpessoas.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    cor = 0
    listpessoas.delete(*listpessoas.get_children())
    vquery = 'SELECT * FROM historico'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[4] == "Pendente":
                listpessoas.insert("", "end", values=i, tags=('red',))
            else:
                listpessoas.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[4] == "Pendente":
                listpessoas.insert("", "end", values=i, tags=('red',))
            else:
                listpessoas.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    calvaloreshistorico()
    ordenarpor = 0

def calvaloreshistorico():
    vquery = 'SELECT SUM(valortotal) FROM historico WHERE status LIKE '+'"Pago"'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        valor = i[0]
    if valor == None:
        valor = 0
    histpagov.configure(text='Valor pago:\n R$:'+str(valor))

    vquery2 = 'SELECT SUM(valortotal) FROM historico WHERE status LIKE '+'"Pendente"'
    linhas2 = Criar.mostrar(vquery2)
    for j in linhas2:
        valor2 = j[0]
    if valor2 == None:
        valor2 = 0
    histpendentev.configure(text='Valor pendente:\n R$:'+str(valor2))

def buscarHistorico():
    global ordenarpor
    listpessoas.column('pessoa', minwidth=288,
                    width=288, anchor=CENTER, stretch=NO)
    listpessoas.column('data', minwidth=100,
                    width=100, anchor=CENTER, stretch=NO)
    listpessoas.column('forma', minwidth=125,
                    width=125, anchor=CENTER, stretch=NO)
    listpessoas.column('valortotal', minwidth=100,
                    width=100, anchor=CENTER, stretch=NO)
    listpessoas.column('status', minwidth=125,
                    width=125, anchor=CENTER, stretch=NO)
    listpessoas.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    cor = 0
    listpessoas.delete(*listpessoas.get_children())
    vquery = "SELECT * FROM historico WHERE nome LIKE '%" + \
        pesquisaprohist.get()+"%'"
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[4] == "Pendente":
                listpessoas.insert("", "end", values=i, tags=('red',))
            else:
                listpessoas.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[4] == "Pendente":
                listpessoas.insert("", "end", values=i, tags=('red',))
            else:
                listpessoas.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    calvaloreshistorico()
    ordenarpor = 0

def mostrarHistorico():
    global ordenarpor
    listpessoas.column('pessoa', minwidth=288,
                    width=288, anchor=CENTER, stretch=NO)
    listpessoas.column('data', minwidth=100,
                    width=100, anchor=CENTER, stretch=NO)
    listpessoas.column('forma', minwidth=125,
                    width=125, anchor=CENTER, stretch=NO)
    listpessoas.column('valortotal', minwidth=100,
                    width=100, anchor=CENTER, stretch=NO)
    listpessoas.column('status', minwidth=125,
                    width=125, anchor=CENTER, stretch=NO)
    listpessoas.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    cor = 0
    listpessoas.delete(*listpessoas.get_children())
    vquery = 'SELECT * FROM historico ORDER BY status DESC'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[4] == "Pendente":
                listpessoas.insert("", "end", values=i, tags=('red',))
            else:
                listpessoas.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[4] == "Pendente":
                listpessoas.insert("", "end", values=i, tags=('red',))
            else:
                listpessoas.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    calvaloreshistorico()
    ordenarpor = 1


    
def altpago():
    global ordenarpor
    try:
        itens = listpessoas.selection()[0]
        valores = listpessoas.item(itens, 'values')
        try:
            Criar.alterastatus("Pago",valores[5])
            if ordenarpor ==1:
                mostrarHistorico()
            else:
                mostarOrdemhis()
        except:
            pass
    except:
        messagebox.showerror(title='Aviso', message='Selecione um item')


def altpen():
    global ordenarpor
    try:
        itens = listpessoas.selection()[0]
        valores = listpessoas.item(itens, 'values')
        try:
            Criar.alterastatus("Pendente",valores[5])
            if ordenarpor ==1:
                mostrarHistorico()
            else:
                mostarOrdemhis()
        except:
            pass
    except:
        messagebox.showerror(title='Aviso', message='Selecione um item')


def detalhespessoa():
    global abrr 
    if abrr == 0:
        threading.Thread(target=detalhespessoa2).start()
        abrr = 1

def detalhespessoa2():
    global abrr
    try:
        itens = listpessoas.selection()[0]
        valores = listpessoas.item(itens, 'values')
        try:
            os.startfile(minhapasta+"\\Meu estoque\\Vendas\\"+valores[5]+".xlsx")
        except:
            pass
    except:
        messagebox.showerror(title='Aviso', message='Selecione um item')
    abrr = 0


def planilhahisto():
    global abrr
    if abrr == 0:
        threading.Thread(target=planilhahisto2).start()
        abrr = 1

def planilhahisto2():
    global abrr
    try:
        os.remove(minhapasta+ '\\Meu estoque\\Clientes.xlsx')
    except:
        pass
    vquery = 'SELECT nome,data,forma,valortotal,status FROM historico'
    linhas = Criar.mostrar(vquery)
    arquivo = xlsxwriter.Workbook(minhapasta+ '\\Meu estoque\\Clientes.xlsx')
    table = arquivo.add_worksheet()
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format.set_align('vcenter')
    currency_format.set_num_format('R$#,##0.00')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format2.set_align('vcenter')
    table.set_column_pixels("A:A",width=188)
    table.set_column_pixels("B:B",width=100)
    table.set_column_pixels("C:C",width=100)
    table.set_column_pixels("D:D",width=120)
    table.set_column_pixels("E:E",width=100)
    table.write(0,0,"Nome do cliente",currency_format2)
    table.write(0,1,"Data",currency_format2)
    table.write(0,2,"Pagamento",currency_format2)
    table.write(0,3,"Valor",currency_format2)
    table.write(0,4,"Status",currency_format2)
    row = 2
    col = 0
    for i in linhas:
        table.write(row,col, i[0],currency_format2)
        table.write(row,col +1, i[1],currency_format2)
        table.write(row,col +2,i[2],currency_format2)
        table.write(row,col +3, i[3],currency_format)
        table.write(row,col +4, i[4],currency_format2)
        row +=1

    arquivo.close()
    os.startfile(minhapasta+"\\Meu estoque\\Clientes.xlsx")
    abrr = 0


def planilhahistoor():
    global abrr
    if abrr == 0:
        threading.Thread(target=planilhahistoor2).start()
        abrr = 1

def planilhahistoor2():
    global abrr
    try:
        os.remove(minhapasta+ '\\Meu estoque\\Clientes.xlsx')
    except:
        pass
    vquery = 'SELECT nome,data,forma,valortotal,status FROM historico ORDER BY status DESC'
    linhas = Criar.mostrar(vquery)
    arquivo = xlsxwriter.Workbook(minhapasta+ '\\Meu estoque\\Clientes.xlsx')
    table = arquivo.add_worksheet()
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format.set_align('vcenter')
    currency_format.set_num_format('R$#,##0.00')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format2.set_align('vcenter')
    table.set_column_pixels("A:A",width=188)
    table.set_column_pixels("B:B",width=100)
    table.set_column_pixels("C:C",width=100)
    table.set_column_pixels("D:D",width=120)
    table.set_column_pixels("E:E",width=100)
    table.write(0,0,"Nome do cliente",currency_format2)
    table.write(0,1,"Data",currency_format2)
    table.write(0,2,"Pagamento",currency_format2)
    table.write(0,3,"Valor",currency_format2)
    table.write(0,4,"Status",currency_format2)
    row = 2
    col = 0
    for i in linhas:
        table.write(row,col, i[0],currency_format2)
        table.write(row,col +1, i[1],currency_format2)
        table.write(row,col +2,i[2],currency_format2)
        table.write(row,col +3, i[3],currency_format)
        table.write(row,col +4, i[4],currency_format2)
        row +=1

    arquivo.close()
    os.startfile(minhapasta+"\\Meu estoque\\Clientes.xlsx")
    abrr = 0




def delonecli():
    global ordenarpor
    try:
        itens = listpessoas.selection()[0]
        valores = listpessoas.item(itens, 'values')
        try:
            m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
            if m == True:
                Criar.deletaprodutohis(valores[5])
            if ordenarpor ==1:
                mostrarHistorico()
            else:
                mostarOrdemhis()
        except:
            pass
    except:
        messagebox.showerror(title='Aviso', message='Selecione um item')



def delallcli():
    global abrr
    if abrr == 0:
        threading.Thread(target=delallcli2).start()
        abrr = 1

def delallcli2():
    global abrr
    global ordenarpor
    try:
        m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
        if m == True:
            Criar.deletatudohis()
        if ordenarpor ==1:
            mostrarHistorico()
        else:
            mostarOrdemhis()
    except:
        pass
    abrr = 0


################ FRAME HISTORICO #################


ambientes = Frame(app, bg='#48d')

Label(ambientes, text='Produtos por ambientes:', bg='#48d', fg='#000',
      font=('arial', 18)).place(x=50, y=9, width=312, height=38)

asomaambien = Label(ambientes, text='', bg='#48d', fg='#000',font=('arial', 15))
asomaambien.place(x=350, y=9, width=220, height=40)

Label(ambientes, text='Novo ambientes:', bg='#48d', fg='#000',
      font=('arial', 18),anchor=W).place(x=900, y=9, width=312, height=38)

Label(ambientes, text='Ambiente:', bg='#48d', fg='#000',
      font=('arial', 18)).place(x=850, y=50, width=130, height=38)

novoambi = Entry(ambientes, width=25, font=('arial', 13))
novoambi.place(x=980, y=50, width=170, height=31)

Label(ambientes, text='Novo produto:', bg='#48d', fg='#000',
      font=('arial', 18),anchor=W).place(x=900, y=250, width=312, height=38)

Label(ambientes, text='Nome:', bg='#48d', fg='#000',
      font=('arial', 18),anchor=E).place(x=800, y=300, width=130, height=38)

nomepronovo = Entry(ambientes, width=25, font=('arial', 13))
nomepronovo.place(x=930, y=300, width=200, height=31)

Label(ambientes, text='QTD:', bg='#48d', fg='#000',
      font=('arial', 18),anchor=E).place(x=800, y=350, width=130, height=38)

qtdpronovo = Entry(ambientes, width=25, font=('arial', 13))
qtdpronovo.place(x=930, y=350, width=200, height=31)

Label(ambientes, text='Valor uni:', bg='#48d', fg='#000',
      font=('arial', 18),anchor=E).place(x=800, y=400, width=130, height=38)

valoepronovo = Entry(ambientes, width=25, font=('arial', 13))
valoepronovo.place(x=930, y=400, width=200, height=31)

produtoambi = ttk.Treeview(ambientes, columns=(
    'qtd','Nome do produto', 'valorunidade','cod','ambi'), show='headings')
produtoambi.column('qtd', minwidth=50,
                width=50, anchor=CENTER, stretch=NO)
produtoambi.column('Nome do produto', minwidth=325,
                width=325, anchor=CENTER, stretch=NO)
produtoambi.column('valorunidade', minwidth=139,
                width=139, anchor=CENTER, stretch=NO)
produtoambi.column('cod', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)
produtoambi.column('ambi', minwidth=1,
                width=1, anchor=CENTER, stretch=NO)

produtoambi.heading('qtd', text='QTD', anchor=CENTER)
produtoambi.heading('Nome do produto',text='Nome do produto', anchor=CENTER)
produtoambi.heading('valorunidade', text='Valor (R$)', anchor=CENTER)
produtoambi.heading('cod',text='cod', anchor=CENTER)
produtoambi.heading('ambi',text='ambi', anchor=CENTER)

produtoambi.tag_configure('oddrow', background="lightblue")
produtoambi.tag_configure('evenrow', background="white")
produtoambi.tag_configure('red', background="orange")

scrollbarhome = Scrollbar(produtoambi)
scrollbarhome.pack(side=RIGHT, fill=Y)
produtoambi.config(yscrollcommand=scrollbarhome.set)
scrollbarhome.config(command=produtoambi.yview)


produtoambi.place(x=50, y=113, width=524, height=575)
  
pesquisaproproambi = Entry(ambientes, width=25, font=('arial', 13))
pesquisaproproambi.place(x=110, y=75, width=300, height=31)



ambien = ttk.Treeview(ambientes, columns=(
    'nome'), show='headings')
ambien.column('nome', minwidth=184,
                width=184, anchor=CENTER, stretch=NO)

ambien.heading('nome', text='Ambientes', anchor=CENTER)

ambien.tag_configure('oddrow', background="lightblue")
ambien.tag_configure('evenrow', background="white")
ambien.tag_configure('red', background="orange")

scrollbarhome = Scrollbar(ambien)
scrollbarhome.pack(side=RIGHT, fill=Y)
ambien.config(yscrollcommand=scrollbarhome.set)
scrollbarhome.config(command=ambien.yview)

                     
ambien.place(x=600, y=70, width=200, height=380)
  
pesquisaambi = Entry(ambientes, width=25, font=('arial', 13))
pesquisaambi.place(x=570, y=30, width=150, height=31)


Label(ambientes, text='Lista para imprimir', bg='#48d', fg='#000',
      font=('arial', 18),anchor=E).place(x=590, y=510, width=210, height=38)

Label(ambientes, text='QTD:', bg='#48d', fg='#000',
      font=('arial', 18),anchor=E).place(x=600, y=610, width=70, height=38)

qtdpraadds = Entry(ambientes, width=25, font=('arial', 13))
qtdpraadds.place(x=670, y=610, width=60, height=38)


def asomav(valor):
    vquery = 'SELECT SUM(valortotal) FROM produto WHERE ambi LIKE '+'"'+valor+'"'
    linhas = Criar.mostrar(vquery)
    for i in linhas: 
        num = i[0]
        if num == None:
            asomaambien.configure(text="")  
        else:
            asomaambien.configure(text="Valor total do ambiente\nR$ "+str(num))  
        

def mostarproambiente():
    try:
        itens = ambien.selection()[0]
        valores = ambien.item(itens, 'values')
        val = valores[0]
    except:
        val = ''
    produtoambi.column('qtd', minwidth=50,
        width=50, anchor=CENTER, stretch=NO)
    produtoambi.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtoambi.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoambi.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    produtoambi.column('ambi', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)                   
    cor = 0
    produtoambi.delete(*produtoambi.get_children())
    vquery = 'SELECT qtuni,nome,valoruni,cod,ambi FROM produto WHERE ambi LIKE '+'"'+val+'"'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[0] == 0:
                produtoambi.insert("", "end", values=i, tags=('red',))
            else:
                produtoambi.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[0] == 0:
                produtoambi.insert("", "end", values=i, tags=('red',))
            else:
                produtoambi.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    asomav(val)
    pesquisaproproambi.delete(0,END)

def buscarproambiente():
    try:
        itens = ambien.selection()[0]
        valores = ambien.item(itens, 'values')
        val = valores[0]
    except:
        val = ''
    produtoambi.column('qtd', minwidth=50,
        width=50, anchor=CENTER, stretch=NO)
    produtoambi.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtoambi.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoambi.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    produtoambi.column('ambi', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)   
    cor = 0
    produtoambi.delete(*produtoambi.get_children())
    vquery = 'SELECT qtuni,nome,valoruni,cod,ambi FROM produto WHERE ambi LIKE '+'"'+val+'" AND nome LIKE \
        '+'"%'+pesquisaproproambi.get()+'%"'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[0] == 0:
                produtoambi.insert("", "end", values=i, tags=('red',))
            else:
                produtoambi.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[0] == 0:
                produtoambi.insert("", "end", values=i, tags=('red',))
            else:
                produtoambi.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    asomav(val)

def buscarproambiente2(event):
    try:
        itens = ambien.selection()[0]
        valores = ambien.item(itens, 'values')
        val = valores[0]
    except:
        val = ''
    produtoambi.column('qtd', minwidth=50,
        width=50, anchor=CENTER, stretch=NO)
    produtoambi.column('Nome do produto', minwidth=325,
                    width=325, anchor=CENTER, stretch=NO)
    produtoambi.column('valorunidade', minwidth=139,
                    width=139, anchor=CENTER, stretch=NO)
    produtoambi.column('cod', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)
    produtoambi.column('ambi', minwidth=1,
                    width=1, anchor=CENTER, stretch=NO)   
    cor = 0
    produtoambi.delete(*produtoambi.get_children())
    vquery = 'SELECT qtuni,nome,valoruni,cod,ambi FROM produto WHERE ambi LIKE '+'"'+val+'" AND nome LIKE \
        '+'"%'+pesquisaproproambi.get()+'%"'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            if i[0] == 0:
                produtoambi.insert("", "end", values=i, tags=('red',))
            else:
                produtoambi.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            if i[0] == 0:
                produtoambi.insert("", "end", values=i, tags=('red',))
            else:
                produtoambi.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    asomav(val)

ambien.bind(sequence="<ButtonRelease-1>",func=buscarproambiente2)

def mostarambients():
    ambien.column('nome', minwidth=184,width=184, anchor=CENTER, stretch=NO)

    cor = 0
    ambien.delete(*ambien.get_children())
    vquery = 'SELECT * FROM ambiente'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            ambien.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            ambien.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    buscarproambiente()

def buscarambientes():
    ambien.column('nome', minwidth=184,width=184, anchor=CENTER, stretch=NO)

    cor = 0
    ambien.delete(*ambien.get_children())
    vquery = 'SELECT * FROM ambiente WHERE nome LIKE '+'"%'+pesquisaambi.get()+'%"'
    linhas = Criar.mostrar(vquery)
    for i in linhas:
        if cor == 0:
            ambien.insert("", "end", values=i, tags=('oddrow',))
            cor = 1
        else:
            ambien.insert("", "end", values=i, tags=('evenrow',))
            cor = 0
    buscarproambiente()

def ev(event):
    buscarambientes()
    
pesquisaambi.bind(sequence="<Return>", func=ev)

def ccambi(event):
    cambiente()

def cambiente():
    threading.Thread(target=cambiente2).start()

def cambiente2():
    nome = novoambi.get()
    if nome == '':
        pass
    else:
        Criar.cadastraambiente(nome.lower())
        jun()
        mostarambients()
        novoambi.delete(0,END)

novoambi.bind(sequence="<Return>",func=ccambi)

def jun():
    vquery = 'SELECT nome,Count(*) FROM ambiente GROUP BY nome HAVING Count(*) > 1' 
    linha = Criar.mostrar(vquery)
    for i in linha:
        Criar.delambiente(i[0])
        Criar.cadastraambiente(i[0])

def aambiente():
    threading.Thread(target=aambiente2).start()

def aambiente2():
    try:
        itens = ambien.selection()[0]
        valores = ambien.item(itens, 'values')
        try:
            nome = novoambi.get()
            if nome == '':
                pass
            else:
                try:
                    Criar.alteambiente(nome.lower(),valores[0])
                    Criar.alteambienteproduto(nome.lower(),valores[0])
                    jun()
                    mostarambients()
                    buscarproambiente()
                    novoambi.delete(0,END)
                except:
                    pass
        except:
            pass
    except:
        messagebox.showwarning(title='Aviso', message='Selecione um item')



def dambiente():
    threading.Thread(target=dambiente2).start()

def dambiente2():
    try:
        itens = ambien.selection()[0]
        valores = ambien.item(itens, 'values')
        try:
            m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
            if m == True:
                Criar.delambiente(valores[0])
                Criar.deletaprodutoambiente(valores[0])
                mostarambients()
                buscarproambiente()
        except:
            pass
    except:
        messagebox.showwarning(title='Aviso', message='Selecione um item')


def cproambiente():  
    try:
        itens = ambien.selection()[0]
        variavels = ambien.item(itens, 'values')
        try:
            a=1
            b = 1.5
            r=0
            nomes = nomepronovo.get()
            qtunis = qtdpronovo.get()
            valores = valoepronovo.get()
            if nomes == '' or qtunis == ''or valores == '':
                messagebox.showwarning(title='ERROR', message='Preencha os dados')
                nomepronovo.focus()
            else:
                try:
                    cod = unidecode(nomes)+ " R$:"+str(valores)
                    img = qrcode.make(cod)
                    img.save('barcode.jpg')
                    valorunis = float(valores.replace(',','.'))
                    type(int(qtunis)) == type(a)
                    type(float(valorunis)) == type(b)
                except:
                    messagebox.showwarning(title='ERROR', message='Preencha os dados corretamente')
                    app.lift()

                if  type(int(qtunis)) == type(a) and type(float(valorunis)) == type(b) and nomes != '':
                        r = 1
                if r == 1:
                    valortotals = int(qtunis) * float(valorunis)
                    me = Criar.cadastraproduto(nomes.lower(),int(qtunis),round(float(valorunis),2),round(float(valortotals),2),cod,int(qtunis),variavels[0])
                    if me != None:
                        messagebox.showerror(title='ERROR', message=me)
                    else:
                        nomepronovo.delete(0,END)
                        qtdpronovo.delete(0,END)
                        valoepronovo.delete(0,END)
                        nomepronovo.focus()
                    mostarproambiente()
        except:
            pass
    except:
        messagebox.showwarning(title='Aviso', message='Selecione o ambiente do produto')

def aproambiente():
    try:
        itens = produtoambi.selection()[0]
        valores = produtoambi.item(itens, 'values')
        try:
            telaAlteraProduto(valores[1],valores[3],valores[4])
        except:
            pass
    except:
        messagebox.showwarning(
            title='ERROR', message='Selecione o item da tabela.')


def dproambiente():
    try:
        itens = produtoambi.selection()[0]
        valores = produtoambi.item(itens, 'values')
        try:
            m = messagebox.askyesno(title='Aviso', message='Deseja deletar.')
            if m == True:
                Criar.deletaproduto(valores[3])
                mostarproambiente()
        except:
            pass
    except:
        messagebox.showwarning(title='Aviso', message='Selecione um item')

    


def addtodospr():
    threading.Thread(target=addtodospr2).start()

def addtodospr2():
    try:
        itens = ambien.selection()[0]
        valores = ambien.item(itens, 'values')
        vquery = "SELECT * FROM produto WHERE ambi LIKE '"+valores[0]+"'"
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            nome = i[0]
            qtuni = i[1]
            valoruni = i[2]
            valortotal = i[3]
            codimg = i[4]
            cod = i[5]
            Criar.cadastraprodutoimpri(nome,qtuni,valoruni,valortotal,codimg,cod)

    except:
        messagebox.showwarning(title='Error',message="Selecione o ambiente")


def addonepro():
    threading.Thread(target=addonepro2).start()

def addonepro2():
    vez = qtdpraadds.get()
    if vez == '':
        vez = 1
    try:
        itens = produtoambi.selection()[0]
        valores = produtoambi.item(itens, 'values')
        vquery = "SELECT * FROM produto WHERE cod LIKE '"+valores[3]+"'"
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            nome = i[0]
            qtuni = i[1]
            valoruni = i[2]
            valortotal = i[3]
            codimg = i[4]
            cod = i[5]
            for j in range(1,int(vez)+1):
                Criar.cadastraprodutoimpri(nome,qtuni,valoruni,valortotal,codimg,cod)
        qtdpraadds.delete(0,END)
        produtoambi.focus()
    except:
        messagebox.showwarning(
            title='ERROR', message='Selecione o produto.')

################ IMAGEM #################             

imagem = PhotoImage(file=minhapasta+"\\1.png")

def criarCoddebarras2OA():
    global abr3
    if abr3 ==0:
        threading.Thread(target=criarCoddebarrasOA).start()
        abr3 = 1

def criarCoddebarrasOA():
    global abr3
    try:
        if os.path.isdir(minhapasta+"\\img"):
            dir = os.listdir(minhapasta + '\\img')
            for file in dir:
                os.remove(minhapasta + '\\img\\'+file)
        else:
            os.mkdir(minhapasta+"\\img")
    except:
        pass
    try:
        os.remove(minhapasta+"\\Meu estoque\\QRCode_Alfabetica.xlsx")
    except:
        pass
    vquery = 'SELECT nome,codimg,valoruni,cod FROM produto ORDER BY nome'
    linhas = Criar.mostrar(vquery)
    try:
        arquivo = xlsxwriter.Workbook(minhapasta+"\\Meu estoque\\QRCode_Alfabetica.xlsx")
    except:
        return
    table = arquivo.add_worksheet()
    table.set_column_pixels("A:C",width=240)
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format3 = arquivo.add_format()
    currency_format3.set_align('center')
    currency_format4 = arquivo.add_format()
    currency_format4.set_align('center')
    currency_format.set_font_size(12)
    currency_format2.set_font_size(10)
    currency_format3.set_font_size(8)
    currency_format4.set_font_size(7)
    row = 1
    row2 = 2
    row3 = 0
    col = 0
    cont = 0
    nun = 1
    
    for i in linhas:
        try:
            img = i[1]
            with open(minhapasta+"\\img\\"+str(nun)+".jpg","wb") as f:
                f.write(img)
            ver = Image.open(minhapasta+"\\img\\"+str(nun)+".jpg")
            sa =ver.resize([200,150],Image.ANTIALIAS)
            sa.save(minhapasta+"\\img\\"+str(nun)+".jpg",quality=100)
            table.set_row_pixels(row,150)
            table.insert_image(row,col,minhapasta+"\\img\\"+str(nun)+".jpg",{'x_offset': 20})
        
            table.set_row_pixels(row2,16)
            table.set_row_pixels(row3,15)
            if int(len(str(i[3]))) <= 24:
                table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format)
            elif int(len(str(i[3]))) <= 28:
                table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format2)
            elif int(len(str(i[3]))) <= 45:
                table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format3)
            elif int(len(str(i[3]))) > 45:
                table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format4)

            if cont == 0:
                frase.configure(text="Carregando.")
            elif cont == 20:
                frase.configure(text="Carregando..")
            elif cont == 40:
                frase.configure(text="Carregando...")
            elif cont == 60:
                cont = 0         
            col +=1

            if col == 3:
                row +=3
                row2 +=3
                row3 +=3
                col =0
                cont +=1

            nun += 1
        except:
            pass

    frase.configure(text="Aguarde")     
    try:   
        arquivo.close() 
        os.startfile(minhapasta+"\\Meu estoque\\QRCode_Alfabetica.xlsx")
    except:
        pass
    frase.configure(text="")
    remover()
    abr3 = 0


def criarCoddebarras2impri2():
    global abr3
    if abr3 ==0:
        threading.Thread(target=criarCoddebarras2impri20).start()
        abr3 = 1

def criarCoddebarras2impri20():
    global abr3
    try:
        if os.path.isdir(minhapasta+"\\img"):
            dir = os.listdir(minhapasta + '\\img')
            for file in dir:
                os.remove(minhapasta + '\\img\\'+file)
        else:
            os.mkdir(minhapasta+"\\img")
    except:
        pass
    try:
        os.remove(minhapasta+"\\Meu estoque\\QRCode_menor.xlsx")
    except:
        pass
    vquery = 'SELECT nome,codimg,valoruni,cod FROM imprimir'
    linhas = Criar.mostrar(vquery)
    try:
        arquivo = xlsxwriter.Workbook(minhapasta+"\\Meu estoque\\QRCode_menor.xlsx")
    except:
        return
    table = arquivo.add_worksheet()
    table.set_column_pixels("A:C",width=240)
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format3 = arquivo.add_format()
    currency_format3.set_align('center')
    currency_format4 = arquivo.add_format()
    currency_format4.set_align('center')
    currency_format.set_font_size(12)
    currency_format2.set_font_size(10)
    currency_format3.set_font_size(8)
    currency_format4.set_font_size(7)
    row = 1
    row2 = 2
    row3 = 0
    col = 0
    cont = 0
    nun = 1
    contsa = 1
    for i in linhas:
        try:
            img = i[1]
            with open(minhapasta+"\\img\\"+str(nun)+".jpg","wb") as f:
                f.write(img)
            ver = Image.open(minhapasta+"\\img\\"+str(nun)+".jpg")
            sa =ver.resize([160,85],Image.ANTIALIAS)
            sa.save(minhapasta+"\\img\\"+str(nun)+".jpg",quality=100)
            table.set_row_pixels(row3,19)
            table.set_row_pixels(row,85)
            table.insert_image(row,col,minhapasta+"\\img\\"+str(nun)+".jpg",{'x_offset': 40,'y_offset':1})
               
            table.set_row_pixels(row2,18)

            if int(len(str(i[3]))) <= 24:
                table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format)
            elif int(len(str(i[3]))) <= 28:
                table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format2)
            elif int(len(str(i[3]))) <= 45:
                table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format3)
            elif int(len(str(i[3]))) > 45:
                table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format4)
            if cont == 0:
                fraseim.configure(text="Carregando.")
            elif cont == 20:
                fraseim.configure(text="Carregando..")
            elif cont == 40:
                fraseim.configure(text="Carregando...")
            elif cont == 60:
                cont = 0          
            col +=1

            if col == 3:
                row +=3
                row2 +=3
                row3 +=3
                col =0
                cont +=1
                contsa += 1

            nun += 1
        except:
            pass

    fraseim.configure(text="Aguarde")   
    try:     
        arquivo.close() 
        os.startfile(minhapasta+"\\Meu estoque\\QRCode_menor.xlsx")
    except:
        pass
    fraseim.configure(text="")
    remover()
    abr3 = 0

def remover():
    try:
        if os.path.isdir(minhapasta+"\\img"):
            dir = os.listdir(minhapasta + '\\img')
            for file in dir:
                os.remove(minhapasta + '\\img\\'+file)
    except:
        pass

def criarCoddebarras2():
    global abr3
    if abr3 ==0:
        threading.Thread(target=criarCoddebarras).start()
        abr3 = 1

def criarCoddebarras():
    global abr3
    try:
        if os.path.isdir(minhapasta+"\\img"):
            dir = os.listdir(minhapasta + '\\img')
            for file in dir:
                os.remove(minhapasta + '\\img\\'+file)
        else:
            os.mkdir(minhapasta+"\\img")
    except:
        pass
    try:
        os.remove(minhapasta+"\\Meu estoque\\QRCode.xlsx")
    except:
        pass
    vquery = 'SELECT nome,codimg,valoruni,cod FROM produto'
    linhas = Criar.mostrar(vquery)
    try:
        arquivo = xlsxwriter.Workbook(minhapasta+"\\Meu estoque\\QRCode.xlsx")
    except:
        return
    table = arquivo.add_worksheet()
    table.set_column_pixels("A:C",width=240)
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format3 = arquivo.add_format()
    currency_format3.set_align('center')
    currency_format4 = arquivo.add_format()
    currency_format4.set_align('center')
    currency_format.set_font_size(12)
    currency_format2.set_font_size(10)
    currency_format3.set_font_size(8)
    currency_format4.set_font_size(7)
    row = 1
    row2 = 2
    row3 = 0
    col = 0
    nun = 1
    cont = 0
    for i in linhas:
        img = i[1]
        with open(minhapasta+"\\img\\"+str(nun)+".jpg","wb") as f:
            f.write(img)
        ver = Image.open(minhapasta+"\\img\\"+str(nun)+".jpg")
        sa =ver.resize([200,150],Image.ANTIALIAS)
        sa.save(minhapasta+"\\img\\"+str(nun)+".jpg",quality=100)
        table.set_row_pixels(row,150)
        table.insert_image(row,col,minhapasta+"\\img\\"+str(nun)+".jpg",{'x_offset': 20})
        table.set_row_pixels(row2,16)

        table.set_row_pixels(row3,15)


        if int(len(str(i[3]))) <= 24:
            table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format)
        elif int(len(str(i[3]))) <= 28:
            table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format2)
        elif int(len(str(i[3]))) <= 45:
            table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format3)
        elif int(len(str(i[3]))) > 45:
            table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format4)
        if cont == 0:
            frase.configure(text="Carregando.")
        elif cont == 20:
            frase.configure(text="Carregando..")
        elif cont == 40:
            frase.configure(text="Carregando...")
        elif cont == 60:
            cont = 0   

        col +=1
        if col == 3:
            row +=3
            row2 +=3
            row3 +=3
            col =0
            cont +=1


        nun += 1

    frase.configure(text="Aguarde")        
    try:
        arquivo.close() 
        os.startfile(minhapasta+"\\Meu estoque\\QRCode.xlsx")
    except:
        pass
    frase.configure(text="")
    remover()
    abr3 = 0

def criarCoddebarrasunico2():
    global abr4
    if abr4 == 0:
        threading.Thread(target=criarCoddebarrasunico).start()
        abr4 = 1
        
def addimpi():
    global abr4
    if abr4 == 0:
        threading.Thread(target=addimpi2).start()
        abr4 = 1

def addimpi2():
    global abr4
    vez = vezess.get()
    if vez == '':
        vez = 1
    try:
        itens = produtos.selection()[0]
        valores = produtos.item(itens, 'values')
        vquery = "SELECT * FROM produto WHERE cod LIKE '"+valores[4]+"'"
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            nome = i[0]
            qtuni = i[1]
            valoruni = i[2]
            valortotal = i[3]
            codimg = i[4]
            cod = i[5]
            for j in range(1,int(vez)+1):
                Criar.cadastraprodutoimpri(nome,qtuni,valoruni,valortotal,codimg,cod)
        vezess.delete(0,END)
        produtos.focus()
        abr4 = 0
    except:
        messagebox.showwarning(
            title='ERROR', message='Selecione o item da tabela.')
        vezess.delete(0,END)
        vezess.focus()
        abr4 = 0
    
def criarCoddebarrasunico():
    global abr4
    vez = vezess.get()
    if vez == '':
        vez = 1
    try:
        itens = produtos.selection()[0]
        valores = produtos.item(itens, 'values')
        try:
            os.remove(minhapasta+"\\Meu estoque\\QRCode_unico\\QRCode_Unico.xlsx")
        except:
            pass
        
        vquery = "SELECT nome,codimg,valoruni,cod FROM produto WHERE cod LIKE '"+valores[4]+"'"
        linhas = Criar.mostrar(vquery)
        try:
            arquivo = xlsxwriter.Workbook(minhapasta+"\\Meu estoque\\QRCode_unico\\QRCode_Unico.xlsx")
        except:
            return
        table = arquivo.add_worksheet()
        table.set_column_pixels("A:C",width=240)
        currency_format = arquivo.add_format()
        currency_format.set_align('center')
        currency_format2 = arquivo.add_format()
        currency_format2.set_align('center')
        currency_format3 = arquivo.add_format()
        currency_format3.set_align('center')
        currency_format4 = arquivo.add_format()
        currency_format4.set_align('center')
        currency_format.set_font_size(12)
        currency_format2.set_font_size(10)
        currency_format3.set_font_size(8)
        currency_format4.set_font_size(7)
        row = 1
        row2 = 2
        row3 = 0
        col = 0
        for i in linhas:
            img = i[1]
            with open("imagem.jpg","wb") as f:
                f.write(img)
            ver = Image.open("imagem.jpg")
            sa =ver.resize([200,150],Image.ANTIALIAS)
            sa.save("imagem.jpg",quality=100)
            for j in range(1,int(vez)+1):
                table.set_row_pixels(row,150)
                table.insert_image(row,col,"imagem.jpg",{'x_offset': 20})
                table.set_row_pixels(row2,16)
                table.set_row_pixels(row3,15)
                if int(len(str(i[3]))) <= 24:
                    table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format)                   
                elif int(len(str(i[3]))) <= 28:
                    table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format2)
                elif int(len(str(i[3]))) <= 45:
                    table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format3)
                elif int(len(str(i[3]))) > 45:
                    table.write(row2,col,str(i[0])+ " R$:"+str(i[2]).replace(".0",""),currency_format4)
                col +=1
                if col == 3:
                    row +=3
                    row2 +=3
                    row3 +=3
                    col =0
        try:
            arquivo.close() 
            os.startfile(minhapasta+"\\Meu estoque\\QRCode_unico\\QRCode_Unico.xlsx")
        except:
            pass
        try:
            os.remove('imagem.jpg')
        except:
            pass
        abr4 = 0
    except:
        messagebox.showwarning(
            title='ERROR', message='Selecione o item da tabela.')
        abr4 = 0

def criarexcelbarras2():
    global abr2
    if abr2 ==0:
        threading.Thread(target=criarexcelbarras).start()
        abr2 = 1

def criarexcelbarras():
    global abr2
    try:
        os.remove(minhapasta+ '\\Meu estoque\\Todos_produtos.xlsx')
    except:
        pass
    vquery = 'SELECT qtuni,nome,valoruni,valortotal,ambi,cod FROM produto'
    linhas = Criar.mostrar(vquery)
    try:
        arquivo = xlsxwriter.Workbook(minhapasta+ '\\Meu estoque\\Todos_produtos.xlsx')
    except:
        return
    table = arquivo.add_worksheet()
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format.set_align('vcenter')
    currency_format.set_num_format('R$ #,##0.00')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format2.set_align('vcenter')
    table.set_column_pixels("A:A",width=50)
    table.set_column_pixels("B:B",width=286)
    table.set_column_pixels("C:D",width=120)
    table.set_column_pixels("E:E",width=100)
    table.set_column_pixels("F:F",width=270)
    table.write(0,0,"QTD",currency_format2)
    table.write(0,1,"Produto",currency_format2)
    table.write(0,2,"Valor uni",currency_format2)
    table.write(0,3,"Valor total",currency_format2)
    table.write(0,4,"Ambiente",currency_format2)
    table.write(0,5,"Cód. do produto",currency_format2)
    row = 1
    col = 0
    for i in linhas:
        table.write(row,col, i[0],currency_format2)
        table.write(row,col +1, i[1],currency_format2)
        table.write(row,col +2,i[2],currency_format)
        table.write(row,col +3, i[3],currency_format)
        table.write(row,col +4 , i[4],currency_format2)
        table.write(row,col +5 , i[5],currency_format2)
        row +=1

    table.write(row,col +3,'=SUM(D2:D'+str(row)+')',currency_format)
    try:
        arquivo.close()
        os.startfile(minhapasta+'\\Meu estoque\\Todos_produtos.xlsx')
    except:
        pass
    abr2 = 0

def criarexcelbarras2OA():
    global abr2
    if abr2 ==0:
        threading.Thread(target=criarexcelbarrasOA).start()
        abr2 = 1

def criarexcelbarrasOA():
    global abr2
    try:
        os.remove(minhapasta+ '\\Meu estoque\\Todos_produtos_Alfabetica.xlsx')
    except:
        pass
    vquery = 'SELECT qtuni,nome,valoruni,valortotal,ambi,cod FROM produto ORDER BY nome'
    linhas = Criar.mostrar(vquery)
    try:
        arquivo = xlsxwriter.Workbook(minhapasta+ '\\Meu estoque\\Todos_produtos_Alfabetica.xlsx')
    except:
        return
    table = arquivo.add_worksheet()
    currency_format = arquivo.add_format()
    currency_format.set_align('center')
    currency_format.set_align('vcenter')
    currency_format.set_num_format('R$ #,##0.00')
    currency_format2 = arquivo.add_format()
    currency_format2.set_align('center')
    currency_format2.set_align('vcenter')
    table.set_column_pixels("A:A",width=50)
    table.set_column_pixels("B:B",width=286)
    table.set_column_pixels("C:D",width=120)
    table.set_column_pixels("E:E",width=100)
    table.set_column_pixels("F:F",width=300)
    table.write(0,0,"QTD",currency_format2)
    table.write(0,1,"Produto",currency_format2)
    table.write(0,2,"Valor uni",currency_format2)
    table.write(0,3,"Valor total",currency_format2)
    table.write(0,4,"Ambiente",currency_format2)
    table.write(0,5,"Cód. do produto",currency_format2)
    
    row = 1
    col = 0
    for i in linhas:
        table.write(row,col, i[0],currency_format2)
        table.write(row,col +1, i[1],currency_format2)
        table.write(row,col +2,i[2],currency_format)
        table.write(row,col +3, i[3],currency_format)
        table.write(row,col +4 , i[4],currency_format2)
        table.write(row,col +5 , i[5],currency_format2)
        
        row +=1

    table.write(row,col +3,'=SUM(D2:D'+str(row)+')',currency_format)
    try:
        arquivo.close()
        os.startfile(minhapasta+'\\Meu estoque\\Todos_produtos_Alfabetica.xlsx')
    except:
        pass
    abr2 = 0

def verpastaa():
    os.startfile(minhapasta+"\\Meu estoque")


def restaura():
    threading.Thread(target=restaura2).start()


def restaura2():
    m = messagebox.askyesno(title='Aviso', message='Deseja restaurar.')
    if m == True:
        vquery = "SELECT cod,qtuni FROM produto"
        linhas = Criar.mostrar(vquery)
        for i in linhas:
            Criar.alteracarctere(i[1],i[0])
        Criar.deletaprodutotudoquase()
        Criar.deletaprodutotudoquase2()
    else:
        pass

def al():
    global vendapo
    if vendapo == 0:
        vendapo = 1
    else:
        vendapo=0
    frameVendas()


################ BOTOES FRAME HOME #################

buscarhome = Button(home, text='Buscar', width=10, command=buscarHome)
buscarhome.configure(bd=1, activebackground="#467", cursor='hand2')
buscarhome.place(x=575, y=75, width=63, height=31)

recarregarHome = Button(home, text='', width=10,command=mostrarProHome )
recarregarHome.configure(bd=1, activebackground="#467",
                     cursor='hand2', image=imagem)
recarregarHome.place(x=650, y=72, width=41, height=38)

ordemAlfabetica = Button(home, text='Ordem alfabética', width=10, command=mostarOrdem)
ordemAlfabetica.configure(bd=1, activebackground="#467", cursor='hand2')
ordemAlfabetica.place(x=50, y=75, width=100, height=31)

NovoHome = Button(home, text='Novo produto', width=10,command=frameAmbientes )
NovoHome.configure(bd=1, activebackground="#467", cursor='hand2')
NovoHome.place(x=775, y=125, width=163, height=38)

AlterarHome = Button(home, text='Alterar produto', width=10, command=alteraProduto)
AlterarHome.configure(bd=1, activebackground="#467", cursor='hand2')
AlterarHome.place(x=963, y=125, width=163, height=38)

DeletarHome = Button(home, text='Deletar produto', width=10, command=deletarpro)
DeletarHome.configure(bd=1, activebackground="#467", cursor='hand2')
DeletarHome.place(x=775, y=200, width=163, height=38)

DeletarHomeTudo = Button(home, text='Deletar tudo', width=10, command=deletarproTudo)
DeletarHomeTudo.configure(bd=1, activebackground="#467", cursor='hand2')
DeletarHomeTudo.place(x=963, y=200, width=163, height=38)

rHomeTudo = Button(home, text='Rec', width=10, command=recproTudo)
rHomeTudo.configure(bd=1, activebackground="#467", cursor='hand2')
rHomeTudo.place(x=1140, y=200, width=38, height=38)

gerarcodeHome = Button(home, text='Gerar QRCode', width=10, command=criarCoddebarras2)
gerarcodeHome.configure(bd=1, activebackground="#467", cursor='hand2')
gerarcodeHome.place(x=775, y=300, width=163, height=38)


gerarxlsxHome = Button(home, text='Gerar planilha', width=10, command=criarexcelbarras2)
gerarxlsxHome.configure(bd=1, activebackground="#467", cursor='hand2')
gerarxlsxHome.place(x=963, y=300, width=163, height=38)

gerarxlsxHomeUnico = Button(home, text='Gerar QRCode único', width=10, command=criarCoddebarrasunico2)
gerarxlsxHomeUnico.configure(bd=1, activebackground="#467", cursor='hand2')
gerarxlsxHomeUnico.place(x=930, y=375, width=188, height=38)

addimprimir = Button(home, text='ADD', width=10, command=addimpi)
addimprimir.configure(bd=1, activebackground="#467", cursor='hand2')
addimprimir.place(x=870, y=375, width=50, height=38)

verPasta = Button(home, text='Ver pasta de arquivos', width=10, command=verpastaa)
verPasta.configure(bd=1, activebackground="#467", cursor='hand2')
verPasta.place(x=775, y=438, width=188, height=38)

carArq = Button(home, text='Carregar arquivo', width=10, command=abrirarq)
carArq.configure(bd=1, activebackground="#467", cursor='hand2')
carArq.place(x=973, y=438, width=188, height=38)

DeletarHomeTudozero = Button(home, text='Deletar produtos com 0 uni.', width=10, command=deletarproTudozero)
DeletarHomeTudozero.configure(bd=1, activebackground="#467", cursor='hand2')
DeletarHomeTudozero.place(x=875, y=538, width=188, height=38)

gerarcodeHomeOA = Button(home, text='Gerar QRCode o. alfabética', width=10, command=criarCoddebarras2OA)
gerarcodeHomeOA.configure(bd=1, activebackground="#467", cursor='hand2')
gerarcodeHomeOA.place(x=875, y=600, width=188, height=38)


gerarxlsxHomeOA = Button(home, text='Gerar planilha o. alfabética', width=10, command=criarexcelbarras2OA)
gerarxlsxHomeOA.configure(bd=1, activebackground="#467", cursor='hand2')
gerarxlsxHomeOA.place(x=875, y=662, width=188, height=38)

################ BOTOES FRAME VENDAS #################   

addProduto = Button(vendas, text='Inserir produto', width=10, command=insertProduto)
addProduto.configure(bd=1, activebackground="#467", cursor='hand2')

alteratudovenda = Button(vendas, text='Alterar', width=10, command=al)
alteratudovenda.configure(bd=1, activebackground="#467", cursor='hand2')

buscarvenda = Button(vendas, text='Buscar', width=10, command=pesquisaProvendas)
buscarvenda.configure(bd=1, activebackground="#467", cursor='hand2')

recarregarvenda = Button(vendas, text='', width=10,command=mostrarProvendas )
recarregarvenda.configure(bd=1, activebackground="#467",
                     cursor='hand2', image=imagem)



cli2 = Button(vendas, text='Cliente 1', width=10, command=incli1)
cli2.configure(bd=1, activebackground="#467", cursor='hand2')
cli2.place(x=100, y=662, width=125, height=38)

cli1 = Button(vendas, text='Cliente 2', width=10, command=incli2)
cli1.configure(bd=1, activebackground="#467", cursor='hand2')
cli1.place(x=300, y=662, width=125, height=38)

delItem = Button(vendas, text='Deletar item', width=10, command=deleteQuase)
delItem.configure(bd=1, activebackground="#467", cursor='hand2')
delItem.place(x=563, y=662, width=125, height=38)

delAll = Button(vendas, text='Deletar tudo', width=10, command=deletAlll)
delAll.configure(bd=1, activebackground="#467", cursor='hand2')
delAll.place(x=725, y=662, width=125, height=38)

avancars = Button(vendas, text='Avançar', width=10, command=frameAvancar)
avancars.configure(bd=1, activebackground="#467", cursor='hand2')
avancars.place(x=1000, y=662, width=125, height=38)

################ BOTOES FRAME AVANCAR #################

voltarAvan = Button(avancar, text='Voltar', width=10, command=frameVendas)
voltarAvan.configure(bd=1, activebackground="#467", cursor='hand2')
voltarAvan.place(x=813, y=662, width=125, height=38)

finalizarAvan = Button(avancar, text='Finalizar', width=10, command=finaliza2)
finalizarAvan.configure(bd=1, activebackground="#467", cursor='hand2')
finalizarAvan.place(x=1000, y=662, width=125, height=38)

rese = Button(avancar, text='Remover desconto', width=10, command=remodesconto)
rese.configure(bd=1, activebackground="#467", cursor='hand2')
rese.place(x=1020, y=100, width=125, height=30)

btvalo = Button(avancar, text='Valor (R$)', width=10, command=calva)
btvalo.configure(bd=1, activebackground="#467", cursor='hand2')
btvalo.place(x=1020, y=200, width=125, height=30)

btpor = Button(avancar, text='Porcentagem (%)', width=10, command=calpor)
btpor.configure(bd=1, activebackground="#467", cursor='hand2')
btpor.place(x=1020, y=240, width=125, height=30)

################ BOTOES FRAME PRODUTOS VENDIDOS ################# 

buscarprodutosVendidos = Button(produtosVendidos, text='Buscar', width=10, command=buscarVend)
buscarprodutosVendidos.configure(bd=1, activebackground="#467", cursor='hand2')
buscarprodutosVendidos.place(x=575, y=75, width=63, height=31)

recarregarProdutoVendidos = Button(produtosVendidos, text='', width=10,command=mostrarVend )
recarregarProdutoVendidos.configure(bd=1, activebackground="#467",
                     cursor='hand2', image=imagem)
recarregarProdutoVendidos.place(x=650, y=72, width=41, height=38)

ordemAlfabeticavend = Button(produtosVendidos, text='Ordem alfabética', width=10, command=mostarOrdemVend)
ordemAlfabeticavend.configure(bd=1, activebackground="#467", cursor='hand2')
ordemAlfabeticavend.place(x=88, y=75, width=125, height=31)

deltaau= Button(produtosVendidos, text='Deletar produto', width=10, command=prodelitem)
deltaau.configure(bd=1, activebackground="#467", cursor='hand2')
deltaau.place(x=875, y=128, width=188, height=38)

deltu= Button(produtosVendidos, text='Deletar Tudo', width=10, command=delintemVendidotudo)
deltu.configure(bd=1, activebackground="#467", cursor='hand2')
deltu.place(x=875, y=188, width=188, height=38)

gerarxlsxVend= Button(produtosVendidos, text='Gerar planilha', width=10, command=criarexcelbarrasVendidos)
gerarxlsxVend.configure(bd=1, activebackground="#467", cursor='hand2')
gerarxlsxVend.place(x=875, y=250, width=188, height=38)

gerarxlsxVendoa= Button(produtosVendidos, text='Gerar planilha o. alfabética', width=10, command=criarexcelbarrasVendidos2OA)
gerarxlsxVendoa.configure(bd=1, activebackground="#467", cursor='hand2')
gerarxlsxVendoa.place(x=875, y=313, width=188, height=38)

verrxlsxVend= Button(produtosVendidos, text='Ver na pasta', width=10, command=verpastaa)
verrxlsxVend.configure(bd=1, activebackground="#467", cursor='hand2')
verrxlsxVend.place(x=875, y=375, width=188, height=38)

rest= Button(produtosVendidos, text='Corrigir estoque ', width=10, command=restaura)
rest.configure(bd=1, activebackground="#467", cursor='hand2')
rest.place(x=875, y=500, width=188, height=38)


################ BOTOES FRAME IMPRIMIR #################

buscarprodutosimpri = Button(impri, text='Buscar', width=10, command=buscarimpri)
buscarprodutosimpri.configure(bd=1, activebackground="#467", cursor='hand2')
buscarprodutosimpri.place(x=575, y=75, width=63, height=31)

recarregarProdutoimpri = Button(impri, text='', width=10,command=mostrarimpri )
recarregarProdutoimpri.configure(bd=1, activebackground="#467",
                     cursor='hand2', image=imagem)
recarregarProdutoimpri.place(x=650, y=72, width=41, height=38)


delimpri= Button(impri, text='Deletar produto', width=10, command=delintemimpri)
delimpri.configure(bd=1, activebackground="#467", cursor='hand2')
delimpri.place(x=775, y=188, width=168, height=38)

delimpritudo= Button(impri, text='Deletar Tudo', width=10, command=delintemimpritudo)
delimpritudo.configure(bd=1, activebackground="#467", cursor='hand2')
delimpritudo.place(x=963, y=188, width=168, height=38)

gerarcodeimpri = Button(impri, text='Gerar QRCode', width=10, command=criarCoddebarras2impri)
gerarcodeimpri.configure(bd=1, activebackground="#467", cursor='hand2')
gerarcodeimpri.place(x=850, y=300, width=200, height=38)

gerarcodeimpri2 = Button(impri, text='Gerar QRCode menor', width=10, command=criarCoddebarras2impri2)
gerarcodeimpri2.configure(bd=1, activebackground="#467", cursor='hand2')
gerarcodeimpri2.place(x=850, y=350, width=200, height=38)

################ BOTOES FRAME HISTORICO #################

buscarprodutoshis = Button(histori, text='Buscar', width=10, command=buscarHistorico)
buscarprodutoshis.configure(bd=1, activebackground="#467", cursor='hand2')
buscarprodutoshis.place(x=575, y=75, width=63, height=31)

recarregarProdutohis = Button(histori, text='', width=10,command=mostarOrdemhis)
recarregarProdutohis.configure(bd=1, activebackground="#467",
                     cursor='hand2', image=imagem)
recarregarProdutohis.place(x=650, y=72, width=41, height=38)

ordemAlfabeticahis = Button(histori, text='Ordenar pendentes', width=10, command=mostrarHistorico)
ordemAlfabeticahis.configure(bd=1, activebackground="#467", cursor='hand2')
ordemAlfabeticahis.place(x=88, y=75, width=125, height=31)

btpago= Button(histori, text='Alterar para pago', width=10, command=altpago)
btpago.configure(bd=1, activebackground="#467", cursor='hand2')
btpago.place(x=815, y=158, width=168, height=38)

btpen= Button(histori, text='Alterar para pendente', width=10, command=altpen)
btpen.configure(bd=1, activebackground="#467", cursor='hand2')
btpen.place(x=1000, y=158, width=168, height=38)

detalhepe = Button(histori, text='Detalhes da compra', width=10, command=detalhespessoa)
detalhepe.configure(bd=1, activebackground="#467", cursor='hand2')
detalhepe.place(x=900, y=250, width=200, height=38)

geplahis = Button(histori, text='Gerar planilha ordenada por status', width=10, command=planilhahistoor)
geplahis.configure(bd=1, activebackground="#467", cursor='hand2')
geplahis.place(x=900, y=300, width=200, height=38)

geplahisor = Button(histori, text='Gerar planilha ordenada por data', width=10, command=planilhahisto)
geplahisor.configure(bd=1, activebackground="#467", cursor='hand2')
geplahisor.place(x=900, y=350, width=200, height=38)

btdelcli= Button(histori, text='Deletar cliente', width=10, command=delonecli)
btdelcli.configure(bd=1, activebackground="#467", cursor='hand2')
btdelcli.place(x=815, y=400, width=168, height=38)

btdelcliall= Button(histori, text='Deletar todos clientes', width=10, command=delallcli)
btdelcliall.configure(bd=1, activebackground="#467", cursor='hand2')
btdelcliall.place(x=1000, y=400, width=168, height=38)

################ BOTOES FRAME Ambientes #################  

buscarprodutosambiente = Button(ambientes, text='Buscar', width=10, command=buscarproambiente)
buscarprodutosambiente.configure(bd=1, activebackground="#467", cursor='hand2')
buscarprodutosambiente.place(x=430, y=75, width=63, height=31)

recarregarProdutoambi = Button(ambientes, text='', width=10,command=mostarproambiente)
recarregarProdutoambi.configure(bd=1, activebackground="#467",
                     cursor='hand2', image=imagem)
recarregarProdutoambi.place(x=500, y=72, width=41, height=38)

buscarambiente = Button(ambientes, text='Buscar', width=10, command=buscarambientes)
buscarambiente.configure(bd=1, activebackground="#467", cursor='hand2')
buscarambiente.place(x=730, y=30, width=63, height=31)

recarregarambientes = Button(ambientes, text='', width=10,command=mostarambients)
recarregarambientes.configure(bd=1, activebackground="#467",
                     cursor='hand2', image=imagem)
recarregarambientes.place(x=800, y=30, width=41, height=38)

cadambiente= Button(ambientes, text='Cadastrar ambiente', width=10, command=cambiente)
cadambiente.configure(bd=1, activebackground="#467", cursor='hand2')
cadambiente.place(x=980, y=100, width=170, height=38)

altambiente= Button(ambientes, text='Alterar ambiente', width=10, command=aambiente)
altambiente.configure(bd=1, activebackground="#467", cursor='hand2')
altambiente.place(x=840, y=160, width=150, height=38)

delambiente= Button(ambientes, text='Deletar ambiente', width=10, command=dambiente)
delambiente.configure(bd=1, activebackground="#467", cursor='hand2')
delambiente.place(x=1000, y=160, width=150, height=38)

cadproambi= Button(ambientes, text='Cadastrar produto', width=10, command=cproambiente)
cadproambi.configure(bd=1, activebackground="#467", cursor='hand2')
cadproambi.place(x=980, y=450, width=170, height=38)

altambiente= Button(ambientes, text='Alterar produto', width=10, command=aproambiente)
altambiente.configure(bd=1, activebackground="#467", cursor='hand2')
altambiente.place(x=840, y=500, width=150, height=38)

delambiente= Button(ambientes, text='Deletar produto', width=10, command=dproambiente)
delambiente.configure(bd=1, activebackground="#467", cursor='hand2')
delambiente.place(x=1000, y=500, width=150, height=38)

addtoproambiente= Button(ambientes, text='Adicionar todos produtos', width=10, command=addtodospr)
addtoproambiente.configure(bd=1, activebackground="#467", cursor='hand2')
addtoproambiente.place(x=600, y=560, width=200, height=38)

addproambiente= Button(ambientes, text='ADD', width=10, command=addonepro)
addproambiente.configure(bd=1, activebackground="#467", cursor='hand2')
addproambiente.place(x=750, y=610, width=50, height=38)

###########################################################     

pesquisapro.focus()
mostrarProHome()
mostrarVend()
deletAlll1()
deletAlll2()
app.bind(sequence="<Return>", func=enter)
app.wm_protocol('WM_DELETE_WINDOW', qui)
app.config(menu=barrademenu)
app.mainloop()