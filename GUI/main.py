import os
from win32com.client import Dispatch, gencache
import pythoncom
import tkinter as tk
from tkinter import IntVar, ttk

index_View = 254
index_Layer = 0
y = (29, 24, 9)
options = ["", "Фадеев", "Жуков", "Пащенко", "Харланов", "Шурпач"]


def getKompasApi():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api7 = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
            module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch
        )
    )
    const = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0)
    return module, api7, const.constants


def get_draw_conteiner(active_doc, module):
    doc2d = module.IKompasDocument2D(active_doc)
    view = get_view(doc2d, module)
    draw_cont = module.IDrawingContainer(view)
    return draw_cont


def get_view(doc2d, module):
    view_manager = doc2d.ViewsAndLayersManager
    Views = view_manager.Views
    if not (Views.View("Штамп")):
        view = Views.Add(1)
        view.X = 0
        view.Y = 0
        view.Number = 255
        view.Name = "Штамп"
        view.Current = True
        view.Update()
    view = Views.View(len(Views) - 1)
    view.Layers.Layer(0).Color = 255
    view.Layers.Layer(0).Update()
    return view


def coord(active_doc):
    format = active_doc.LayoutSheets.Item(0).Format
    x = format.FormatWidth - 180
    y = 75
    return x, y


def save_doc(active_doc):
    path = f"{active_doc.Path}\\Чертежи в pdf\\{active_doc.Name}"
    path = path.split(".")[0]
    if save_pdf.get():
        active_doc.SaveAs(f"{path}.pdf")
    else:
        active_doc.Save()
    # active_doc.Close(1)
    print("Document saved successfully.")


def set_stamp_hol(module, active_doc, doc2d):
    draw_cont = get_draw_conteiner(active_doc, module)
    ins_manager = module.IInsertionsManager(active_doc)
    ins_definition = ins_manager.AddDefinition(
        0, "Штамп", rf"\\server\\Чтение и запись\\Обмен файлами\\VOL.frw"
    )
    x, y = coord(active_doc)
    ins = draw_cont.InsertionObjects
    ins = ins.Add(ins_definition)
    ins.SetPlacement(x, y, 0, False)
    ins.Update()


def set_text_stamp(active_doc, case_text, id_stamp):
    for i in range(3):
        format = active_doc.LayoutSheets.Item(0)
        text = format.Stamp.Text(id_stamp[i])
        text.Clear()
        textLine = text.Add()
        textItem = textLine.Add()
        textItem.Str = case_text[i]
        print(text)
        format.Stamp.Update()
        print(format.Stamp.Text(id_stamp[i]).Str)


def set_sign(active_doc, case_text, id_stamp, module):
    for i in range(3):
        format = active_doc.LayoutSheets.Item(0)
        text = format.Stamp.Text(id_stamp[i]).Str
        if text != "":
            signs_path = r"\\server\\Чтение и запись\\Подписи\\" + text + ".png"
            view_manager = module.IKompasDocument2D(active_doc).ViewsAndLayersManager
            Views = view_manager.Views
            view = Views.Add(1)
            view.X = 0
            view.Y = 0
            view.Number = 254
            view.Name = "Подписи"
            view.Current = True
            view.Update()
            view = Views.View("Подписи")
            img_view = module.IDrawingContainer(view).Rasters
            img_view = img_view.Add()
            img_view.SetPlacement(format.Format.FormatWidth - 150, y[i], 0, False)
            img_view.FileName = signs_path
            img_view.Scale = 0.045
            img_view.Update()
            print(signs_path)
            print(img_view)


def main():
    case_text = (combo1.get(), combo2.get(), combo3.get())
    module, api7, const = getKompasApi()
    app = api7.Application
    app.HideMessage = 1
    if app:
        for i in range(app.Documents.Count):
            app.Documents(i).Active = True
            active_doc = app.Documents(i)
            doc2d = module.IKompasDocument2D(active_doc)
            view = get_view(doc2d, module)
            if view.ObjectCount < 1:
                set_stamp_hol(module, active_doc, doc2d)
            else:
                obj = module.IDrawingContainer(view).Objects(
                    0
                )  # Все объекты на системном
                for i in range(len(obj)):
                    obj[i].Delete()  # Удаляем
                set_stamp_hol(module, active_doc, doc2d)
            id_stamp = (110, 111, 114)
            if skip.get():
                set_text_stamp(active_doc, case_text, id_stamp)
            set_sign(active_doc, case_text, id_stamp, module)
            save_doc(active_doc)
            view_manager = module.IKompasDocument2D(active_doc).ViewsAndLayersManager
            Views = view_manager.Views
            view = Views.View("Подписи")
            if view:
                view.Delete()
            app.HideMessage = 0


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Любимый ХОЛСИМ")
    root.geometry("400x300")
    skip = tk.BooleanVar(value=False)
    save_pdf = tk.BooleanVar(value=False)
    combo1 = ttk.Combobox(values=options)
    combo1.pack()
    combo2 = ttk.Combobox(values=options)
    combo2.pack()
    combo3 = ttk.Combobox(values=options)
    combo3.pack()
    check_button = tk.Checkbutton(
        text="Не заполнять штамп", variable=skip, onvalue=False, offvalue=True
    )
    check_button.pack()
    check_button2 = tk.Checkbutton(
        text="Сохранять pdf", variable=save_pdf, onvalue=True, offvalue=False
    )
    check_button2.pack()
    btn = tk.Button(text="Заполнить", command=main)
    btn.pack()
    root.mainloop()
