from kivy.app import App
from kivy.graphics import Color, Rectangle
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.button import Button
import openpyxl

class IncApp(App):
    def build(self):
        layout = BoxLayout(orientation='vertical', padding=100)
        layout.bind(size=self._update_rect, pos=self._update_rect)

        with layout.canvas.before:
            self.bg = Rectangle(source='teste.jpg', pos=layout.pos, size=layout.size)

        self.inc_label = Label(text='Número do INC:', font_size=24, bold=True)
        self.inc_input = TextInput(font_size=20, size_hint_y=None, height=40)
        layout.add_widget(self.inc_label)
        layout.add_widget(self.inc_input)

        self.problem_label = Label(text='Nível do problema:', font_size=24, bold=True)
        problem_layout = BoxLayout(orientation='horizontal')
        self.problem_buttons = []
        for level in ['Leve', 'Médio', 'Alto', 'Crítico']:
            button = ToggleButton(text=level, group='problem_level')
            button.background_color = (0.5, 0.5, 1, 1)
            self.problem_buttons.append(button)
            problem_layout.add_widget(button)
        layout.add_widget(self.problem_label)
        layout.add_widget(problem_layout)

        self.name_label = Label(text='Nome da pessoa:', font_size=24, bold=True)
        self.name_input = TextInput(font_size=20, size_hint_y=None, height=40)
        layout.add_widget(self.name_label)
        layout.add_widget(self.name_input)

        self.type_label = Label(text='Qual Setor:', font_size=24, bold=True)
        type_layout = BoxLayout(orientation='horizontal')
        self.type_buttons = []
        for inc_type in ['DAT', 'CD']:
            button = ToggleButton(text=inc_type, group='inc_type')
            button.background_color = (0.5, 0.5, 1, 1)
            self.type_buttons.append(button)
            type_layout.add_widget(button)
        layout.add_widget(self.type_label)
        layout.add_widget(type_layout)

        self.submit_button = Button(text='Enviar')
        self.submit_button.bind(on_press=self.submit)
        layout.add_widget(self.submit_button)

        return layout

    def _update_rect(self, instance, value):
        self.bg.pos = instance.pos
        self.bg.size = instance.size

    def submit(self, instance):
        inc_number = self.inc_input.text
        problem_level = None
        for button in self.problem_buttons:
            if button.state == 'down':
                problem_level = button.text
                break
        name = self.name_input.text
        inc_type = None
        for button in self.type_buttons:
            if button.state == 'down':
                inc_type = button.text
                break

        wb = openpyxl.load_workbook('inc.xlsx')
        sheet = wb.active
        sheet.append([inc_number, problem_level, name, inc_type])
        wb.save('inc.xlsx')

        self.inc_input.text = ''
        for button in self.problem_buttons:
            button.state = 'normal'
        self.name_input.text = ''
        for button in self.type_buttons:
            button.state = 'normal'

IncApp().run()