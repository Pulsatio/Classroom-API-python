from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import googleapiclient.errors as errors
import simplejson

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile



# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/classroom.courses', 'https://www.googleapis.com/auth/classroom.coursework.students',
          'https://www.googleapis.com/auth/classroom.rosters','https://www.googleapis.com/auth/classroom.announcements',
          'https://www.googleapis.com/auth/classroom.topics','https://www.googleapis.com/auth/classroom.profile.emails']


IDCLASE = {}
enrolCodeCLASE = {}

INICIAL = ['3 AÑOS','4 AÑOS','5 AÑOS']
PRIMARIA = ['1ro','1ro-1','1ro-2','2do','2do-1','2do-2','3ro','3ro-1','3ro-2','4to','4to-1','4to-2','5to','5to-1','5to-2','5to-3','6to','6to-1','6to-2']
SECUNDARIA = ['1roSec','1roSec-1','1roSec-2','1roSec-3','1roSec-4',
                '2doSec','2doSec-1','2doSec-2','2doSec-3','2doSec-4',
                '3roSec','3roSec-1','3roSec-2','3roSec-3','3roSec-4',
                '4toSec','4toSec-1','4toSec-2','4toSec-3','4toSec-4','4toSec-5',
                '5toSec','5toSec-1','5toSec-2','5toSec-3']


SELECCION = ['SEL 1','SEL 2','SEL 3','SEL 4','SEL 5']
ESPECIAL = ['ESPECIAL 1','ESPECIAL 2']
OLIMPICOS = ['OLIMPICOS 1','OLIMPICOS 2','OLIMPICOS 3']

def pause():
    input('Oprima enter para continuar')

def mapearClases(servicio,cantidadClases=500):
    output = servicio.courses().list(pageSize=cantidadClases).execute()
    clases = output.get('courses', [])  
    if not clases:
        print('No se encontraron clases.')
    else:
        for clase in clases:
            IDCLASE[ clase['name'] ] = clase['id']
            try :
                enrolCodeCLASE[ clase['name'] ] = clase['enrollmentCode']
            except:
                continue
def listarClases(servicio, cantidadClases = 500):
    # Se imprime lista de 10(valor por defecto) primeros clases
    output = servicio.courses().list(pageSize=cantidadClases).execute()
    clases = output.get('courses', [])  
    if not clases:
        print('No se encontraron clases.')
    else:
        print('Los clases son los siguientes:')
        for clase in clases:
            IDCLASE[clase['name']] = clase['id']
            print(clase['name'], clase['id'])
    
def crearClase(servicio,nombre):
    # Se crea un classroom con el nombre dado
    datos = {
        'name': nombre,
        'ownerId': 'me',
        'courseState': 'ACTIVE'
    }
    clase = servicio.courses().create(body=datos).execute()
    idClase= clase.get('id')
    print('Clase creada: {0} ({1})'.format(clase.get('name'), idClase ))
    IDCLASE[nombre] = idClase
    enrolCodeCLASE[nombre] = clase.get('enrollmentCode')


def eliminarClase(servicio,nombre,idClase):
    servicio.courses().delete(id=idClase).execute()
    print('Clase eliminada con exito')
    del IDCLASE[nombre]
    del enrolCodeCLASE[nombre]


def agregarTopicoaClase(servicio,idClase, nombreTopico):
    datos = {
        "name": nombreTopico
    }
    topico = servicio.courses().topics().create(courseId=idClase, body = datos).execute()
    print('Topico creado: ', topico['name'] , topico['topicId'])
    return topico

def agregarTareaaClase(servicio,idClase,idTopico,tituloTarea, tipoTarea):
    datos = {
        'title': tituloTarea,
        'workType': tipoTarea,
        'topicId': idTopico,
        'state': 'PUBLISHED',
    }
    tarea = servicio.courses().courseWork().create(courseId=idClase, body=datos).execute()
    print('Tarea creada con ID {0}'.format(tarea.get('id')))
    return

def agregarProfesoraClase(servicio, emailProfesor, idClase):
    datos = {
        'userId': emailProfesor
    }
    try:
        teacher = servicio.courses().teachers().create(courseId=idClase,
                                                      body=datos).execute()
        print('Se agrego al profesor {0} a la clase con ID "{1}"'.format(teacher.get('profile').get('name').get('fullName'),idClase))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 409):
            print ('El profesor con correo "{0}" ya es parte de la clase'.format(emailProfesor))
        else:
            raise
    return

def eliminarProfesordeClase(servicio,emailProfesor,idClase):
    try:
        teacher = servicio.courses().teachers().create(courseId=idClase,userId=emailProfesor).execute()
        print('Se elimino al profesor con correo {0} de la clase con ID "{1}"'.format(emailProfesor,idClase))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 404):
            print ('El profesor con correo "{0}" ya es parte de la clase'.format(emailProfesor))
        else:
            raise
    return

def invitarPersonaaClase(servicio,emailAlumno, idClase, tipo ):

    datos = {
        'userId': emailAlumno,
        'role': tipo,
        'courseId': idClase
    }
    try:
        persona = servicio.invitations().create(body=datos).execute()
        print('{0} con correo {1} fue invitado a la clase con ID "{1}"'.format(tipo , persona.get('userId'), idClase))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 409):
            print ('{0} con correo "{1}" ya fue invitado a la clase'.format(tipo,emailAlumno))

    return


def eliminarAlumnodeClase(servicio,emailAlumno,idClase,codigoClase):
    try:
        student = servicio.courses().students().delete(courseId=idClase,userId=emailAlumno).execute()
        print('El alumno con correo {0} se elimino clase con ID "{1}"'.format(emailAlumno,idClase))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 404):
            print('El alumno no se encuentra matriculado en la clase.')
        else:
            print('Error de correo.')


def eliminarAlumnos(servicio,archivo,clase,idClase,enrolCode):
    correos=leerTXT(archivo)
    cont=1
    for correo in correos:
        print("Linea ",end="")
        print(cont,end="")
        print(" : ")
        eliminarAlumnodeClase(servicio,correo,idClase,enrolCode)
        cont+=1



def agregarAlumnoaClase(servicio, emailAlumno, idClase,codigoClase):
    datos = {
        'userId': emailAlumno
    }
    try:
        student = servicio.courses().students().create(courseId=idClase,enrollmentCode=codigoClase,body=datos).execute()
        print('Alumno {0} esta ahora cursando la clase con ID "{1}"'.format(student.get('profile').get('name').get('fullName'),idClase))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 409):
            print('El alumno ya se encuentra matriculado en la clase.')
        else:
            print('Error de correo.')

def obtenerClaseporID(servicio, idClase):
    # obtener la clase mediante su id
    return servicio.courses().get(id=idClase).execute()

def obtenerCodigoClase(servicio,idClase):
    # obtener el codigo de la clase para agregar invasimente
    return obtenerClaseporID(servicio,idClase)['enrollmentCode']
def leerTXT(archivo):
    f = open (archivo,'r')
    data =  f.read().split('\n') 
    return data



def registrarAlumnos(servicio,archivo,clase,idClase,enrolCode):
    correos=leerTXT(archivo)
    cont=1
    for correo in correos:
        print("Linea ",end="")
        print(cont,end="")
        print(" : ")
        agregarAlumnoaClase(servicio,correo,idClase,enrolCode)
        cont+=1

def obtenerDatosAlumno(servicio,idAlumno):
    try:
        profile = servicio.userProfiles().get(userId=idAlumno).execute()
        return profile
    except:
        print('Error en el id del Alumno')    



def creacionMasiva( servicio, listaClases, listaTopicos , listaTareas):

    for clase in listaClases:
        claseActual = crearClase(servicio, clase)
        for topico in listaTopicos:
            nuevoTopico = agregarTopicoaClase(servicio,claseActual['id'] , topico)
            for tarea in listaTareas:
                agregarTareaaClase(servicio,claseActual['id'],nuevoTopico['topicId'],  tarea, 'ASSIGNMENT')

    return





def listarAlumnosEnClase(servicio,clase,idClase):
    cant=0
    data = servicio.courses().students().list(courseId=idClase).execute()
    if 'students' in data:
        alumnos = data['students']
    else :
        print('No se encontraron alumnos en la clase')
        return
    for alu in alumnos:
        cant+=1
        print(alu)
        print(cant,end=" :")
        print(alu['profile']['name']['fullName'])
    try:
        pToken = data['nextPageToken']
    except:
        pToken = None
#    print(pToken)
    while pToken!=None:
        data = servicio.courses().students().list(courseId=idClase,pageToken=pToken).execute()
        alumnos = data['students']
        for alu in alumnos:
            cant+=1
            print(cant,end=" :")
            print(alu['profile']['name']['fullName'])
        try:
            pToken = data['nextPageToken']
        except :
            break





def menu(servicio):
    while 1 :
        print()
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print("¿Que desea realizar?")
        print("1. Listar las clases subidas a classroom")
        print("2. Agregar alumnos a una clase")
        print("3. Eliminar alumnos de una clase")
        print("4. Agregar profesor a una clase")
        print("5. Eliminar profesor de una clase")
        print("6. Crear clase")
        print("7. Eliminar clase")
        print("8. Obtener informacion de una clase")
        print("9. Obtener informacion de un alumno")
        print("0. Salir del programa")
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print()

        op = input()

        if op=='1':
            listarClases(servicio)
            print()
            pause();
        elif op=='2':
            menu2(servicio)
        elif op=='3':
            menu3(servicio)
        elif op=='4':
            clase = input()
            if clase in IDCLASE:
                emailProfesor = input()
                agregarProfesoraClase(servicio,emailProfesor,IDCLASE[clase])
            else:
                print('No existe esa clase')
                print('Intentelo nuevamente')
            print()
            pause();    
        elif op=='5':
            clase = input()
            if clase in IDCLASE:
                emailProfesor = input()
                eliminarProfesordeClase(servicio,emailProfesor,IDCLASE[clase])
            else:
                print('No existe esa clase')
                print('Intentelo nuevamente')
            print()
            pause();                  
       
        elif op=='6':
            print('Ingrese el nombre de la clase a crear')
            clase = input()
            if clase in IDCLASE:
                print('Existe una clase con ese nombre')
            else:
                crearClase(servicio,clase);
            print()
            pause();                   
        elif op=='7':
            print('Ingrese el nombre de la clase que va a eliminar')
            clase = input()
            if clase in IDCLASE:
                idClase = IDCLASE[clase]
                eliminarClase(servicio,clase,idClase);
            else:
                print('No existe una clase con ese nombre')
            print()
            pause();
        elif op=='8':
            print("Ingrese el nombre de la clase")
            clase = input()
            if clase in IDCLASE:
                menu8(servicio,clase,IDCLASE[clase])
            else:
                print('No existe una clase con ese nombre')
        elif op=='9':
            print("En construccion")
            print()
            pause()
#            menu9(servicio)
        elif op=='0':
            break;
        else :
            print("Opcion invalida")
            print("Ingrese nuevamente la opcion")
            print()
            pause();


def menu2(servicio):
    while 1 :
        print()
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print("Ingrese una de las opciones")
        print("1. Agregar alumno a una clase")
        print("2. Agregar un lista de alumnos a una clase")
        print("0. Volver al menu anterior")
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print()
        op  = input()
        if op=='1':
            print("Ingrese el nombre de la clase")
            nombreClase = input()
            if nombreClase in IDCLASE:
                print("Ingrese el correo del alumno")
                correo = input()
                idClase = IDCLASE[nombreClase] 
                enrolCode = enrolCodeCLASE[nombreClase]
                agregarAlumnoaClase(servicio,correo,idClase,enrolCode)
            else:
                print("No existe esa Clase")
                print("Intentelo nuevamente")
            print()
            pause();
        elif op=='2':
            print("Ingrese el nombre de la clase")
            nombreClase = input()
            if nombreClase in IDCLASE:
                print("Ingrese el nombre del archivo txt que contiene los correos de los alumnos")
                print("El archivo debe estar en la misma carpeta que en la del programa")
                archivo = input()
                idClase = IDCLASE[nombreClase] 
                enrolCode = enrolCodeCLASE[nombreClase]
                try:
                    registrarAlumnos(servicio,archivo,idClase,enrolCode)
                except:
                    print('Error. Intentelo nuevamente')
            else:
                print("No existe esa Clase")
                print("Intentelo nuevamente")
            print()
            pause();
        elif op=='0':
            break;
        else:
            print("Opcion invalida")
            print("Ingrese nuevamente la opcion")
            print()
            pause(); 

def menu3(servicio):
    while 1 :
        print()
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print("Ingrese una de las opciones")
        print("1. Eliminar alumno de una clase")
        print("2. Eliminar un lista de alumnos de una clase")
        print("0. Volver al menu anterior")
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print()
        op  = input()
        if op=='1':
            print("Ingrese el nombre de la clase")
            nombreClase = input()
            if nombreClase in IDCLASE:
                print("Ingrese el correo del alumno")
                correo = input()
                idClase = IDCLASE[nombreClase] 
                enrolCode = enrolCodeCLASE[nombreClase]
                eliminarAlumnodeClase(servicio,correo,idClase,enrolCode)
            else:
                print("No existe esa Clase")
                print("Intentelo nuevamente")
            print()
            pause();
        elif op=='2':
            print("Ingrese el nombre de la clase")
            nombreClase = input()
            if nombreClase in IDCLASE:
                print("Ingrese el nombre del archivo txt que contiene los correos de los alumnos")
                print("El archivo debe estar en la misma carpeta que en la del programa")
                archivo = input()
                idClase = IDCLASE[nombreClase] 
                enrolCode = enrolCodeCLASE[nombreClase]
                try:
                    eliminarAlumnos(servicio,archivo,idClase,enrolCode)
                except:
                    print('Error. Intentelo nuevamente')
            else:
                print("No existe esa Clase")
                print("Intentelo nuevamente")
            print()
            pause();

        elif op=='0':
            break;
        else:
            print("Opcion invalida")
            print("Ingrese nuevamente la opcion") 
            print()
            pause();

def menu8(servicio,clase,idClase):
    while 1 :
        print()
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print("Clase '{0}' con id:'{1}'",format(clase, idClase))
        print("Escoja una opcion")
        print("1. Listar los alumnos de la clase")
        print("2. Encontrar un alumno")
        print("3. Listar los profesores de la clase")
        print("4. Encontrar un profesor")
        print("5. Listar los cursos de la clase")
        print("0. Volver al menu anterior")
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
        print()
        op = input()
        if op=='1':
            listarAlumnosEnClase(servicio,clase,idClase)
            pause()
            print()
        elif op=='2':
            pause()
            print()
        elif op=='3':
            pause()
            print()
        elif op=='4':
            pause()
            print()
        elif op=='5':
            pause()
            print()
        elif op=='0':
            break
        else:
            print("Opcion invalida")
            print("Intentelo nuevamente")
            pause()
            print()


def menu9(servicio):
    while 1 :
        print()
        print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")

def main():
    #Obteniendo las credenciales
    creds = None

    #En token.pickle guardamos el acceso del usuario y lo reutilizamos, se crea en la primera ejecución

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # Si no hay credenciales validas el usuario debe loguearse

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    #Usamos el servicio con las credenciales obtenidas
    servicio = build('classroom', 'v1', credentials=creds)

    mapearClases(servicio)
    
    menu(servicio)

if __name__ == '__main__':
    main()
