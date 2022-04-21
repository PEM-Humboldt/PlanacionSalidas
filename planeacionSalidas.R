install.packages("googlesheets4")
install.packages("mailR")
install.packages("rjava")
install.packages("devtools", dep = T)
install_github("rpremrajGit/mailR")
install.packages("gmailr")
install.packages("googledrive")

library(googledrive)
library(mailR)
library(googlesheets4)
library(mailR)
library(gmailr)
library(devtools)
library(rJava)

####Auntenfificación de cuentas vinculo R - google

#######Ingresar a:  https://console.developers.google.com/project. desde la cuenta gmail a vincular.
## Buscar el proyecto "PalneacionSalidas" y con calidad de editor cree su propio ID de clientes OAuth 2.0

##descargue y asignele una carpeta al archivo "client secret" y copie la clave de la API, remplace y corra la siguiente linea

#dirección de la carpeta asiganada.

path2 <- "C:/Users/hector.arango/Documents/Humboldt/Fibras/Programación salidas de campo/client_secret_199019144146-jbb36q8gitghul9sm76v3vs9n6klv360.apps.googleusercontent.com (1).json"
gm_auth_configure(key="AIzaSyCmwXAvyu3yo35BzFaSkfce9ZEKNQ9fX_k",secret="client_secret_199019144146-jbb36q8gitghul9sm76v3vs9n6klv360.apps.googleusercontent.com",path = path2)
gm_auth_configure(key="AIzaSyCmwXAvyu3yo35BzFaSkfce9ZEKNQ9fX_k",secret="199019144146-35q4kll4m0nfk84lglggun64spsguuf1.apps.googleusercontent.com",path = path)

### Usar archivo Cliensecret

setwd("C:/Users/hector.arango/Documents/Humboldt/Fibras/Programación salidas de campo")

#use_secret_file("PROJ-NAME.json")
#use_secret_file("harangoudea.json")

gmailr::use_secret_file('client_secret_199019144146-jbb36q8gitghul9sm76v3vs9n6klv360.apps.googleusercontent.com (1).json')

path <- "C:/Users/hector.arango/Documents/Humboldt/Fibras/Programación salidas de campo/PROJ-NAME.json"

path2 <- "C:/Users/hector.arango/Documents/Humboldt/Fibras/Programación salidas de campo/client_secret_199019144146-jbb36q8gitghul9sm76v3vs9n6klv360.apps.googleusercontent.com (1).json"
gm_auth_configure(key="AIzaSyCmwXAvyu3yo35BzFaSkfce9ZEKNQ9fX_k",secret="client_secret_199019144146-jbb36q8gitghul9sm76v3vs9n6klv360.apps.googleusercontent.com",path = path2)

gm_auth_configure(path=path)
########## cargar hojas especificas de googe sheet

#Confirmar Email
gs4_auth(email="harango@humboldt.org.co")
gs4_user()

u=read_sheet("https://docs.google.com/spreadsheets/d/1N3TMsmkVGhSseSjRpXipxppB5NcEpuCy76pAT0j8huU/edit#gid=1758415088")

ss=as_id("https://docs.google.com/spreadsheets/d/1SDeTjCn5iX_M_ixQEtoDvrn4MTNZxSKv4SobtWnD2_c/edit#gid=242943464")
Scp=read_sheet(ss="https://docs.google.com/spreadsheets/d/1SDeTjCn5iX_M_ixQEtoDvrn4MTNZxSKv4SobtWnD2_c/edit#gid=242943464",range="Planeacion")
Scdp=read_sheet(ss="https://docs.google.com/spreadsheets/d/1SDeTjCn5iX_M_ixQEtoDvrn4MTNZxSKv4SobtWnD2_c/edit#gid=242943464",range="Datos Personales")


######## Lista de correos investigadores

nmail=vector()
for(i in 1:dim(Scdp[4])[1]){
  nmail[i]=as.character(Scdp[3][i,])
}



########### Actualización Tablas Anticipos y solicitud vuelos
Solc=function(Scp,Scdp){

for(i in 1:dim(Scp)[1]){

  if(Scp$Confirmacion[i]=="si"){

### gastos de Viaje
idin=Scp$`Correos electronicos pasajeros`[i]
idin=paste(c(strsplit(idin,","))[[1]],"@humboldt.org.co",sep="")
ninv=match(idin,nmail)

if(Scp$`Agendas iguales`[i]=="si"){
Ta=strsplit(Scp$Gastos_de_Viaje[i],",")[[1]]
data1=data.frame(Scp$Salida[i],"NA",Scdp$`Nombre completo`[ninv],Ta[1],Ta[2],Ta[3],Ta[4],Ta[5],Ta[6],Ta[7])
sheet_append(ss, data1, sheet ="Relacion GV")

}else{
  Ta=apply(MARGIN=1,as.data.frame(strsplit(Scp$Gastos_de_Viaje[i][[1]],"_")[[1]]),FUN=function(x){strsplit(x,",")})
  taok=list()
  for(j in 1:length(Ta)){
    taok[j]=Ta[[j]]
  }
  taok=do.call(rbind,taok)
  taok=as.data.frame(cbind(Scp$Salida[i],"NA",Scdp$`Nombre completo`[ninv],taok))
  sheet_append(ss, taok, sheet ="Relacion GV")
  }
  }else{
    next
  }


###Anticipos

  if(Scp$Anticipos_concepto[[i]]=="no"){next}
  else{
datan=match(Scp$`Responsable Anticipo 1`[i],Scdp$`Nombre completo`)

data2=data.frame(Scp$Salida[i],Scp$No_Solicitud[[i]],Scp$Anticipos_concepto[[i]],Scdp$`Nombre completo`[datan],Scdp$Banco[datan],Scdp$Cuenta[datan],Scdp$Número[[datan]],
                 Scp$Bienes[i],Scp$`Materiales por comprar`[i],Scp$Envios[i],Scp$Entradas[i],Scp$Rubro_Alimentación[[i]],Scp$`Servicios Personales`[i],
                 Scp$Transporte[i],Scp$`Gasolina Peajes`[i],Scp$Arrendamiento[[i]],Scp$`Inscripcion Capacitaciones`[i])


sheet_append(ss,data2, sheet ="Anticipo")
}


#### Transporte interno

t1=apply(MARGIN=1,as.data.frame(strsplit(Scp$Nvehiculos_recorridos[i][[1]],"_")[[1]]),FUN=function(x){strsplit(x,",")})
t1ok=list()
for(j in 1:length(t1)){
  t1ok[j]=t1[[j]]
}
t1ok=do.call(rbind,t1ok)

t1ok=as.data.frame(cbind(Scp$Salida[i],Scp$Tipo_de_transporte_Terrestre[i],t1ok,Scp$Responsable[i],Scdp$Teléfono...9[match(Scp$Responsable[i],Scdp$`Nombre completo`)]))

sheet_append(ss,t1ok, sheet ="Relacion Transporte")

}

}

######### Envío de Correo de notificación
Cornt=function(Scp,Scdp,emailfrom){

nmail=vector()
for(i in 1:dim(Scdp[4])[1]){
  nmail[i]=as.character(Scdp[3][i,])
}


for(i in 1:dim(Scp)[1]){

  if(Scp$Confirmacion[i]=="ej"){
    next}
  if(Scp$Confirmacion[i]=="no"){
    next
  }else{

    rmid=paste(Scp$`Correo Investigadores`[i],Scp$`Correos logistica`[i],sep=",")
    rmind=strsplit(rmid,",")
    rmind=paste(rmind[[1]],"@humboldt.org.co",sep="")
    a=rmind[1]
    for(x in 2:length(rmind)){
      a=paste(a,rmind[x],sep=",")
    }

    test<- gm_mime(

      To = a,
      From = emailfrom,


      Subject = paste(Scp$Salida[i],Scp$Componente[i],Scp$Objeto_solicitud[i]),
      body = paste("Buenos días, a continuación, se describen las necesidades para realizar la siguiente  salida de campo

    1. descripción general:

    Proyecto:",Scp$Proyecto[i],"
    Componente:",Scp$Componente[i],"
    Objetivo de la actividad:",Scp$Objeto_solicitud[i],"
    Responsable de la actividad:",Scp$Responsable[i],"
    Participantes:",Scp$Asistentes[i],"
    Lugar:",Scp$Municipio[i],"
    Inicio:",Scp$Start_Date[i],"
    Fin:",Scp$End_Date[i],"


    Requerimientos para la Actividad

    2. Transporte y Gastos de viaje

Aéreo y gastos de viaje:

https://docs.google.com/spreadsheets/d/1SDeTjCn5iX_M_ixQEtoDvrn4MTNZxSKv4SobtWnD2_c/edit#gid=1649448113

Terrestre desplazamiento interno:

https://docs.google.com/spreadsheets/d/1SDeTjCn5iX_M_ixQEtoDvrn4MTNZxSKv4SobtWnD2_c/edit#gid=1243454053


    CAR:",Scp$CAR_Transporte[i],"

    3.Plan de movilización de Ecopetrol:  (Solo es necesario para ingresar a predios de ecopetrol)
    De acuerdo con los procedimientos de Ecopetrol esta información se debe enviar diez (10) días
    calendario antes de la salida, por lo tanto, se solicita confirmar que el formato está debidamente
    diligenciado a Héctor Manuel Arango (harango@humboldt.org.co), para que lo valide y  remita a
               Ricardo Adolfo Villarraga Danderino, quien es el encargado de enviar esta información a Ecopetrol.

               Se dispone de una carpeta en Drive donde reposará la información de cada salida de campo
               incluyendo este formato (en el mensaje se enviará el link al formato correspondiente)

              link: https://docs.google.com/spreadsheets/d/1a2gNkxzKaYpfjtpCMtDkEAkVIomivAYz/edit#gid=856034162
    4. Guías:",Scp$Guia[i],"

    5. Materiales y equipos requeridos:",Scp$Materiales[i],"

    6. Anticipos:",Scp$Anticipos_concepto[i],"

       Valor Anticipo:",Scp$Valor_Total_Anticipo[i],"
               Imprevistos del 10%:",Scp$Requiere_imprevisto_del_10[i]
                   ,sep=" "))

    gmailr::gm_send_message(
      test,
      "multipart",
      thread_id = NULL,
      user_id = "me")


  }
}

}




Solc(Scp,Scdp)
Cornt(Scp,Scdp,emailfrom)

