from openpyxl import*
from gurobipy import*

#conjuntos
K= [i for i in range(1,7)] #numeros
R = [i for i in range(1,15)] #regiones
C1 = [(i,j)for i in K for j in K] #celdas


#parametros 
#OPERACIONES
O={1:"SUMA",2:"SUMA",3:"SUMA",4:"DIVISION",5:"RESTA",6:"RESTA",7:"DIVISION",8:"RESTA",
   9:"SUMA",10:"DIVISION",11:"RESTA",12:"SUMA",13:"SUMA",14:"RESTA"}
#RESULTADOS_ESPERADOS
E={1:6,2:12,3:16,4:3,5:5,6:3,7:2,8:1,9:13,10:3,11:1,12:16,13:11,14:3}

#CELDAS POR REGION
C={1:[(1,1),(1,2),(2,1)],2:[(3,1),(3,2),(3,3)],3:[(4,2),(4,1),(5,1),(6,1)],
   4:[(5,2),(6,2)],5:[(5,3),(4,3)],6:[(2,2),(2,3)],7:[(1,3),(1,4)],8:[(2,4),(3,4)],
   9:[(6,3),(6,4),(5,4)],10:[(6,6),(6,5)],11:[(5,6),(5,5)],12:[(4,6),(4,5),(4,4),(3,5)],
   13:[(2,6),(3,6),(2,5)],14:[(1,6),(1,5)]}


#Modelo de Optimización
m = Model("Punto2")

#Definicion de variables de decisión
x = m.addVars(K,C1,vtype=GRB.BINARY, name = "x")
y = m.addVars(C1,vtype=GRB.INTEGER, lb=1, ub=6, name="y")
z1 = m.addVars(R,vtype=GRB.BINARY, name = "z1")
z2 = m.addVars(R,vtype=GRB.BINARY, name = "z2")


#Restricciones del problema
#solo un numero en una casilla
for p in C1:
    m.addConstr(quicksum(x[k,p[0],p[1]] for k in K) == 1)
#no repetir numeros un filas 
for k in K:
    for j in K:
        m.addConstr(quicksum(x[k,i,j] for i in K) == 1)
#no repetir nuemeros en columnas 
for k in K:
    for i in K:
        m.addConstr(quicksum(x[k,i,j] for j in K) == 1)
#restriccion para equiparar variables binarias y continuas        
for p in C1:
    m.addConstr(y[p] == quicksum(k * x[k, p[0],p[1]] for k in K))
        
#restriccion para la suma 
for r in R:
    if O[r] == "SUMA": 
        m.addConstr(quicksum(y[p] for p in C[r]) == E[r])
#restriccion para la resta
for r in R:
    if O[r] == "RESTA":
        for p in C[r]:
            for p2 in C[r]:
                if (p2[0] == p[0]+1 and p2[1] == p[1]) or (p2[0] == p[0] and p2[1] == p[1]+1): #como hay 2 inidces sobre el mismo conjunto solo hay 2 casos que intereasn 
                    m.addConstr(y[p] -y[p2] == E[r] -2*E[r]*z1[r])
#restriccion para la division 
for r in R:
    if O[r] == "DIVISION":
        for p in C[r]:
            for p2 in C[r]:
                if (p2[0] == p[0]+1 and p2[1] == p[1]) or (p2[0] == p[0] and p2[1] == p[1]+1):
                    #caso1 y[p]>y[p2]
                    m.addConstr(y[p] >= E[r]*y[p2] -1e6*z2[r])
                    m.addConstr(y[p] <= E[r]*y[p2] +1e6*z2[r])
                    #caso2 y[p2]>y[p]
                    m.addConstr(y[p2] >= E[r]*y[p] -1e6*(1-z2[r]))
                    m.addConstr(y[p2] <= E[r]*y[p] +1e6*(1-z2[r]))

#Funcion Objetivo
m.setObjective(quicksum(y[i,j] for i in K for j in K),GRB.MAXIMIZE)

#
m.update()
m.setParam("Outputflag",0)
m.optimize()
z = m.getObjective().getValue()

#Imrpimir resultados en consola
print(z)

for p in C1:
    print(y[p[0],p[1]])