import sqlite3
from sqlite3 import Error
import os
from tkinter import messagebox


minhapasta = os.path.dirname(__file__)

if os.path.isdir(minhapasta + '\\Meu estoque'):
    pass
else:
    os.mkdir(minhapasta + '\\Meu estoque')

def conect():
    con=None
    try:
        con = sqlite3.connect(minhapasta + '\\Meu estoque\\DataBase\\DataBase.db') 
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)
    return con


def cria():
    if os.path.isdir(minhapasta + '\\Meu estoque\\DataBase'):
        pass
    else:
        os.mkdir(minhapasta + '\\Meu estoque\\DataBase')
    vcon = conect()
    c = vcon.cursor()

    c.execute('CREATE TABLE IF NOT EXISTS produto (nome  TEXT , qtuni integer, valoruni real, \
        valortotal real, codimg BLOB NOT NULL , cod text PRIMARY KEY,caracteres integer,ambi TEXT)')

    c.execute('CREATE TABLE IF NOT EXISTS produtosvendidos (nome  TEXT , qtuni integer, valoruni real, \
        valortotal real,cod text PRIMARY KEY,codes text)')

    c.execute('CREATE TABLE IF NOT EXISTS produtosquasevendidos (nome  TEXT , qtuni integer, valoruni real, \
        valortotal real, cod text PRIMARY KEY,codes text, FOREIGN KEY(codes) REFERENCES produto(cod) )')

    c.execute('CREATE TABLE IF NOT EXISTS produtosquasevendidos2 (nome  TEXT , qtuni integer, valoruni real, \
        valortotal real, cod text PRIMARY KEY,codes text, FOREIGN KEY(codes) REFERENCES produto(cod) )')

    c.execute('CREATE TABLE IF NOT EXISTS produtore (nome  TEXT , qtuni integer, valoruni real, \
        valortotal real, codimg BLOB NOT NULL , cod text PRIMARY KEY,caracteres integer,ambi TEXT)')
    
    c.execute('CREATE TABLE IF NOT EXISTS produto2 (nome  TEXT , qtuni integer, valoruni real, \
        valortotal real,cod text ,caracteres integer,ambi TEXT)')

    c.execute('CREATE TABLE IF NOT EXISTS imprimir (nome  TEXT , qtuni integer, valoruni real, \
        valortotal real, codimg BLOB NOT NULL , cod text)')

    c.execute('CREATE TABLE IF NOT EXISTS historico (nome  TEXT , data TEXT, forma TEXT, \
        valortotal real, status TEXT , cod text)')

    c.execute('CREATE TABLE IF NOT EXISTS ambiente (nome  TEXT)')

def convertToBinaryData(filename):
    with open(filename, 'rb') as file:
        blobData = file.read()
    return blobData

def cadastraambiente(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO ambiente (nome) VALUES (?) ',(nome,))
        vcon.commit()
    except Error as ex:
        return ex

def delambiente(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM ambiente WHERE nome = '+'"'+nome+'"')
        vcon.commit()
    except Error as ex:
        return ex

def alteambiente(nomes,nomea):
    nome = nomes 
    nomeant = nomea
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('UPDATE ambiente SET nome = ? WHERE nome = ?',(nome,nomeant))
        vcon.commit()
    except Error as ex:
        return ex



def cadastrahist(nomes,datas,formas,valortotals,statuss,cods):
    nome = nomes 
    data = datas
    forma = formas
    valortotal = valortotals
    status = statuss
    cod = cods
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO historico (nome,data,forma,valortotal,status,cod) VALUES (?,?,?,?,?,?)\
            ',(nome,data,forma,valortotal,status,cod))
        vcon.commit()
    except Error as ex:
        return ex

def alterastatus(nomes,cods):
    nome = nomes 
    cod = cods
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('UPDATE historico SET status = ? WHERE cod = ?',(nome,cod))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        return me


def deletaprodutohis(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM historico WHERE cod = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)  

def deletatudohis():
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM historico')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)  

def cadastraprodutore(nomes,qtunis,valorunis,valortotals,codim,cods,caracteress,ambis):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    codimg = codim
    cod = cods
    caracteres = caracteress
    ambi = ambis
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO produtore (nome,qtuni,valoruni,valortotal,codimg,cod,caracteres,ambi\
            ) VALUES (?,?,?,?,?,?,?,?)',(nome,qtuni,valoruni,valortotal,codimg,cod,caracteres,ambi))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        return me

def cadastraproduto2(nomes,qtunis,valorunis,valortotals,cods,caracteress,ambis):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    cod = cods
    caracteres = caracteress
    ambi = ambis
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO produto2 (nome,qtuni,valoruni,valortotal,cod,caracteres,ambi) VALUES (?,?,?,?,?,?,?)\
            ',(nome,qtuni,valoruni,valortotal,cod,caracteres,ambi))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        return me

def deleproduto2(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produto2 WHERE cod = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   

def deleTudoPro2():
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produto2')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   

def cadastraproduto(nomes,qtunis,valorunis,valortotals,cods,caracteress,am):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    codimg = (minhapasta + '\\barcode.jpg')
    cod = cods
    caracteres = caracteress
    ambi = am
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO produto (nome,qtuni,valoruni,valortotal,codimg,cod,caracteres,ambi) VALUES (?,?,?,?,?,?,?,?)\
            ',(nome,qtuni,valoruni,valortotal,convertToBinaryData(codimg),cod,caracteres,ambi))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        if str(ex) == 'UNIQUE constraint failed: produto.cod':
            me = 'Este produto já está cadastrado'
        return me

def cadastraprodutoimpri(nomes,qtunis,valorunis,valortotals,img,cods):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    codimg = img
    cod = cods
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO imprimir (nome,qtuni,valoruni,valortotal,codimg,cod) VALUES (?,?,?,?,?,?)',(nome,qtuni,valoruni,valortotal,codimg,cod))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        if str(ex) == 'UNIQUE constraint failed: produto.cod':
            me = 'Este produto já está cadastrado'
        return me

def cadastraprodo2(nomes,qtunis,valorunis,valortotals,codim,cods,caracteress,ambis):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    codimg = codim
    cod = cods
    caracteres = caracteress
    ambi = ambis
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO produto (nome,qtuni,valoruni,valortotal,codimg,cod,caracteres,ambi) VALUES (?,?,?,?,?,?,?,?)\
            ',(nome,qtuni,valoruni,valortotal,codimg,cod,caracteres,ambi))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        return me


def alteracarctere(nomes,cods):
    nome = nomes 
    cod = cods
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('UPDATE produto SET caracteres = ? WHERE cod = ?',(nome,cod))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        return me

def vender(qtd,valort,cods):
    quant = qtd 
    valor = valort
    cod = cods
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('UPDATE produto SET qtuni =? ,valortotal=? WHERE cod = ?',(quant,valor,cod))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        return me


def alteraproduto(nomes,qtunis,valorunis,valortotals,cods,antnome,caracteress,ambis):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    codimg = (minhapasta + '\\barcode.jpg')
    cod = cods
    ant = antnome
    caracteres=caracteress
    ambi = ambis
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('UPDATE produto SET nome = ?, qtuni = ?, valoruni = ? \
            , valortotal=?, codimg = ? ,cod = ? ,caracteres =?,ambi = ? \
                WHERE cod = ?',(nome,qtuni,valoruni,valortotal,convertToBinaryData(codimg),cod,caracteres,ambi,ant))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        if str(ex) == 'UNIQUE constraint failed: produto.cod':
            me = 'Este produto já foi cadastrado'
        return me

def alteambienteproduto(nomes,nomea):
    nome = nomes 
    nomeant = nomea
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('UPDATE produto SET ambi = ? WHERE ambi = ?',(nome,nomeant))
        vcon.commit()
    except Error as ex:
        return ex

def deleTudoPro():
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produto')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)     

def deleTudoProrec():
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtore')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)  

def deltud():
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosvendidos')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)     

def deltudimpri():
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM imprimir')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)    

def deletaprodutoambiente(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produto WHERE ambi = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)  

def deletaproduto(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produto WHERE cod = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   
    
def deletaprodutoquasevv(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosquasevendidos WHERE codes = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   
    
def deletaprodutoquasevv2(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosquasevendidos2 WHERE codes = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex) 

def mostrar(query):
    vcon = conect()
    c = vcon.cursor()
    c.execute(query)
    res = c.fetchall()
    vcon.close()
    return res


def cadastraprodutoquase(nomes,qtunis,valorunis,valortotals,cods,codess):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    cod = cods
    codes = codess
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO produtosquasevendidos (nome,qtuni,valoruni,valortotal,cod,codes) VALUES (?,?,?,?,?,?)',(nome,qtuni,valoruni,valortotal,cod,codes))
        vcon.commit()
    except Error as ex:
        return ex

def cadastraprodutoquase2(nomes,qtunis,valorunis,valortotals,cods,codess):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    cod = cods
    codes = codess
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO produtosquasevendidos2 (nome,qtuni,valoruni,valortotal,cod,codes) VALUES (?,?,?,?,?,?)',(nome,qtuni,valoruni,valortotal,cod,codes))
        vcon.commit()
    except Error as ex:
        return ex


      
def deletaprodutoquasevend(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosquasevendidos WHERE cod = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   


def deletaprodutoquasevend20(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosquasevendidos WHERE codes = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)  

def deletaprodutoquasevend40(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosquasevendidos2 WHERE codes = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex) 

def deletaprodutoquasevend2(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosquasevendidos2 WHERE cod = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)  

def deletaprodutotudoquase():
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosquasevendidos')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   

def deletaprodutotudoquase2():
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosquasevendidos2')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   

def cadastraprodutoVendido(nomes,qtunis,valorunis,valortotals,cods,codess):
    nome = nomes 
    qtuni = qtunis
    valoruni = valorunis
    valortotal = valortotals
    cod = cods
    codes = codess
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('INSERT INTO produtosvendidos (nome,qtuni,valoruni,valortotal,cod,codes) VALUES (?,?,?,?,?,?)',(nome,qtuni,valoruni,valortotal,cod,codes))
        vcon.commit()
    except Error as ex:
        me = str(ex)
        return me

def deletaprodutovend(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosvendidos WHERE cod = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   

def deletaprodutoimpri(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM imprimir WHERE cod = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)  

def deletaprodutovend2(nomes):
    nome = nomes 
    vcon = conect()
    c = vcon.cursor()
    try:
        c.execute('DELETE FROM produtosvendidos WHERE codes = '+'"'+nome+'"')
        vcon.commit()
        vcon.close()
    except Error as ex:
        messagebox.showwarning(title='ERROR', message=ex)   