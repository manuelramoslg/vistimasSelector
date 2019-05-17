# pip3 install xlrd
# pip3 install pandas
# pip3 install asciimatics

from __future__ import division
from asciimatics.scene import Scene
from asciimatics.screen import Screen
from asciimatics.effects import RandomNoise
from asciimatics.effects import Cycle, Stars
from asciimatics.renderers import FigletText, Rainbow
from asciimatics.exceptions import ResizeScreenError

import os
import sys
import random
import pandas as pd
from pandas import ExcelFile

def get_arguments():
    if len(sys.argv) < 3: # Falta validar si los argumentos disponibles
        print("************************************************")
        print("Ingrese Argumentos:")
        print("Generaciones disponibles: (g1, g2, g3, g4, g5)")
        print("Modos disponibles: (normal, suspenso)")
        print("Ejemplo: python3 vistima.py g5 suspenso")
        print("Para salir presione la tecla 'X'")
        print("************************************************\n")
        sys.exit()
    else:
        generation = sys.argv[1]
        mode = sys.argv[2]
        return generation, mode

def get_student_from_xlsx(generation):
    cwd = os.getcwd()
    df = pd.read_excel(cwd + "/estudiantes.xlsx")
    students = df[(generation).capitalize()].values
    return random.choice(students)

def suspense(screen):
    # Leer xlsx y seleccionar un estudiante
    student = get_student_from_xlsx(generation)
    student_split = student.split()
    text = "{} {} \n{} {}".format(
        student_split[0],
        student_split[1],
        student_split[2],
        student_split[3]
    )
    effects = [
        RandomNoise(screen,
                    signal=Rainbow(screen,
                                   FigletText(text)))
    ]
    screen.play([Scene(effects, -1)], stop_on_resize=True)

def normal(screen):
    # Leer xlsx y seleccionar un estudiante
    student = get_student_from_xlsx(generation)
    student_split = student.split()
    effects = [
        Cycle(
            screen,
            FigletText("{} {}".format(student_split[0], student_split[1]), font='big'),
            screen.height // 2 - 8),
        Cycle(
            screen,
            FigletText("{} {}".format(student_split[2], student_split[3]), font='big'),
            screen.height // 2 + 3),
        Stars(screen, (screen.width + screen.height) // 2)
    ]
    screen.play([Scene(effects, 500)])

while True:
    try:
        # Obtener argumentos
        arg = get_arguments()
        generation = arg[0]
        mode = arg[1]
        # seleccionar efecto
        if mode == "normal":
            Screen.wrapper(normal)
        else:
            Screen.wrapper(suspense)
        sys.exit(0)
    except ResizeScreenError:
        pass
