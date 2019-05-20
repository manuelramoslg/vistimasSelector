# pip3 install xlrd
# pip3 install pandas
# pip3 install asciimatics

from __future__ import division
from asciimatics.scene import Scene
from asciimatics.screen import Screen
from asciimatics.effects import RandomNoise
from asciimatics.effects import Cycle, Stars
from asciimatics.renderers import FigletText, Rainbow, SpeechBubble, Rainbow
from asciimatics.exceptions import ResizeScreenError
from asciimatics.effects import Stars, Print
from asciimatics.particles import RingFirework, SerpentFirework, StarFirework, PalmFirework
from random import randint, choice
import random
import sys

import os
import pandas as pd
from pandas import ExcelFile

def get_arguments():
    if len(sys.argv) < 3: # Falta validar si los argumentos disponibles
        print("************************************************")
        print("Ingrese Argumentos:")
        print("Generaciones disponibles: (g1, g2, g3, g4, g5)")
        print("Modos disponibles: (normal, suspenso, fuego_artificial)")
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

def fireworks(screen):
    scenes = []
    # Leer xlsx y seleccionar un estudiante
    student = get_student_from_xlsx(generation)
    student_split = student.split()
    effects = [
        Stars(screen, screen.width),
        Print(screen,
              SpeechBubble("Pulsa espacio para volver a ver"),
              y=screen.height - 3,
              start_frame=300)
    ]
    for _ in range(20):
        fireworks = [
            (PalmFirework, 25, 30),
            (PalmFirework, 25, 30),
            (StarFirework, 25, 35),
            (StarFirework, 25, 35),
            (StarFirework, 25, 35),
            (RingFirework, 20, 30),
            (SerpentFirework, 30, 35),
        ]
        firework, start, stop = choice(fireworks)
        effects.insert(
            1,
            firework(screen,
                     randint(0, screen.width),
                     randint(screen.height // 8, screen.height * 3 // 4),
                     randint(start, stop),
                     start_frame=randint(0, 250)))

    effects.append(Print(screen,
                         Rainbow(screen, FigletText("{} {}".format(student_split[0], student_split[1]), font='big')),
                         screen.height // 2 - 6,
                         speed=1,
                         start_frame=100))
    effects.append(Print(screen,
                         Rainbow(screen, FigletText("{} {}".format(student_split[2], student_split[3]), font='big')),
                         screen.height // 2 + 1,
                         speed=1,
                         start_frame=100))
    scenes.append(Scene(effects, -1))

    screen.play(scenes, stop_on_resize=True)

while True:
    try:
        # Obtener argumentos
        arg = get_arguments()
        generation = arg[0]
        mode = arg[1]
        # seleccionar efecto
        if mode == "normal":
            Screen.wrapper(normal)
        elif mode == "fuego_artificial":
            Screen.wrapper(fireworks)
        else:
            Screen.wrapper(suspense)
        sys.exit(0)
    except ResizeScreenError:
        pass
