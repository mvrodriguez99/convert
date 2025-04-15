from fastapi import FastAPI
from pydantic import BaseModel
from conversor_prueba import cantidad_a_numero  # importa tu función

app = FastAPI()

class Entrada(BaseModel):
    cantidad_letra: str

@app.post("/convertir/")
def convertir(entrada: Entrada):
    numero = cantidad_a_numero(entrada.cantidad_letra)
    return {"cantidad": numero}