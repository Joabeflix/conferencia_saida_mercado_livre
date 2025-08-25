import tkinter as tk


class Interface:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.geometry('100x100')
        self.botao = tk.Button(self.root, text="Teste", command=lambda: self.teste())
        self.botao.pack()
        self.entreada_teste = tk.Entry(self.root).pack()
        self.root.mainloop()

    def teste(self):
        print('Exec')
        self.root.geometry("200x200")
        # self.root.mainloop()        




app = Interface()

