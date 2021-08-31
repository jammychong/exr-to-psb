# install pywin32 module
# python -m pip install pywin32
import sys
import win32com.client
import os


render_pass_keywords = ["albedo", "background", "coat", "direct", "indirect", "diffuse", "opacity",
        "sheen", "specular", "sss", "transmission"]

# TODO Research if there are other folders where exr_io plugin could be located
def exr_io_check():
    paths = ("C:\\Program Files\\Adobe", "C:\\Program Files (x86)\\Adobe")
    for path in paths:
        if os.path.exists(path):
            for folder in os.listdir(path):
                if "photoshop" in folder.lower():
                    exr_io_path = path + "\\" + folder + "\\Plug-ins\\Exr-IO.8bi"
                    if os.path.exists(exr_io_path):
                        return True
    return False


def ps_connect():
    try:
        psApp = win32com.client.GetActiveObject("Photoshop.Application")
    except:
        psApp = win32com.client.Dispatch("Photoshop.Application")
    return psApp


def open_exrio(psApp, exr_file, apply_alpha):
    desc1 = win32com.client.Dispatch('Photoshop.ActionDescriptor')
    desc2 = win32com.client.Dispatch('Photoshop.ActionDescriptor')
    desc1.PutBoolean(psApp.StringIDToTypeID('dontRecord'), True)
    desc1.PutBoolean(psApp.StringIDToTypeID('forceNotify'), False)
    desc1.PutPath(psApp.CharIDToTypeID('null'), exr_file)

    # Erx-Io argunent:
    # ioal = Add Alpha to all layers
    # iocm = cryptomatte Layers

    settings = ("ioty", "iosa", "ioac", "ioal", "iocm", "ioca", "iocd", "ioll", "ioci", "iodw", "iocg", "iosr")
    values = (True, True, False, apply_alpha, True, False, False, False, True, True, True, True)

    settings_values = list(zip(settings, values))

    for item in settings_values:
        desc2.PutBoolean(psApp.CharIDToTypeID(item[0]), item[1])
    desc2.PutInteger(psApp.CharIDToTypeID("iocw"), 100)

    desc1.PutObject(psApp.CharIDToTypeID('As  '), psApp.StringIDToTypeID('3d-io Exr-IO'), desc2)
    psApp.ExecuteAction(psApp.CharIDToTypeID('Opn '), desc1, 3)


def get_document(psApp):
    doc = psApp.Application.ActiveDocument
    return doc

def render_layers_visibility(render_layers, value):
    for layer in render_layers:
        layer.visible = value


def save_psb(psApp, filename):
    desc1 = win32com.client.Dispatch('Photoshop.ActionDescriptor')
    desc2 = win32com.client.Dispatch('Photoshop.ActionDescriptor')
    desc1.PutBoolean(psApp.StringIDToTypeID('maximizeCompatibility'), True)
    desc2.PutObject(psApp.CharIDToTypeID('As  '), psApp.CharIDToTypeID('Pht8'), desc1)
    desc2.PutPath(psApp.CharIDToTypeID('In  '), filename)
    desc2.PutBoolean(psApp.CharIDToTypeID('LwCs'), True)
    psApp.ExecuteAction(psApp.CharIDToTypeID('save'), desc2, 3)


def psb_name(exr_file):
    psb_file = exr_file[:-4] + ".psb"
    return psb_file


def change_bit_detph(psApp, bit_depth):
    desc1 = win32com.client.Dispatch('Photoshop.ActionDescriptor')
    desc1.PutInteger(psApp.CharIDToTypeID('Dpth'), bit_depth)
    desc1.PutBoolean(psApp.CharIDToTypeID('Mrge'), False)
    psApp.ExecuteAction(psApp.CharIDToTypeID('CnvM'), desc1, 3)


def close_application(psApp):
    docs = psApp.Application.Documents
    if len(docs) == 1:
        psApp.Quit()


def copy_file_contents_to_clipboard(psApp, path):
    open_exrio(psApp, path, False)
    doc = get_document(psApp)
    doc.layers[0].Copy()
    doc.Close(2)


def create_layer_from_file(psApp, doc, layer_name, path):
    copy_file_contents_to_clipboard(psApp, path)
    print("{} copied successfully".format(layer_name))
    psApp.activeDocument = doc
    doc.Paste()
    layer = doc.activeLayer
    layer.name = str(layer_name)
       

def read_crypto_elements(exr_file):
    data = []
    with open(exr_file, 'rb') as f:
        crypto_objects = []
        crypto_layers = 0
        for line in f.readlines():
            str_line = str(line)
            line_data = str_line.split("\\")

            for string in line_data:
                if string[3:5] == '{"':
                    # print(string)
                    # print("-------------------")
                    dictionary = string.split('"}')
                    dictionary_list = dictionary[0][5:].split('","')
                    
                    for i in range(1, len(dictionary_list), 2):
                        obj = dictionary_list[i].split('":"')
                        crypto_objects.append(obj[0])
                    crypto_layers += 1
                    if crypto_layers == 5:
                        return crypto_objects
    return crypto_objects


def add_crypto_layers(psApp, doc, crypto_file):
    open_exrio(psApp, crypto_file, True)
    crypto_doc = get_document(psApp)
    crypto_layers = [layer for layer in crypto_doc.Layers]
    crypto_layers.sort(key=lambda e: e.name)
    
    
    for layer in crypto_layers:
        psApp.activeDocument = crypto_doc
        layer.Copy()
        crypto_name = layer.Name
        crypto_name = crypto_name.split(".")[-1]
        psApp.activeDocument = doc
        doc.Paste()
        new_layer = doc.activeLayer 
        red_channel = doc.Channels.Item('Red')
        doc.Selection.Load(red_channel)
        new_channel = doc.Channels.Add()
        new_channel.name = crypto_name
        doc.Selection.Store(new_channel)
        doc.Selection.Deselect()
        new_layer.Delete()

    crypto_doc.Close(2)


def create_multexr_psb(exr_files, output, crypto_state, bit_value):
    # for _file in exr_files:
    #     print(_file["path"])
    # return "None"
    beauty_index = 0
    crypto_layers = []
    base_passes = {"name": "Render BASE", "layers":[]}
    render_passes = {"name": "Render Passes", "layers":[]}
    aovs_passes = {"name": "AOV_Passes", "layers":[]}
    lighting_passes = {"name": "Lighting Passes", "layers":[]}

    for i in range(len(exr_files)):
        if "beauty" in exr_files[i]["name"]:
            beauty_index = i
            break

    psApp = ps_connect()
    open_exrio(psApp, exr_files[beauty_index]["path"], True)
    main_doc = get_document(psApp)

    for i in range(len(exr_files)):
        if i != beauty_index:
            if "crypto" in exr_files[i]["name"]:
                if crypto_state:
                    crypto_layers.append(exr_files[i]["path"])
            else:
                print(exr_files[i]["path"])
                create_layer_from_file(psApp, main_doc, exr_files[i]["name"], exr_files[i]["path"])

    if len(crypto_layers) > 0:
        for crypto_file in crypto_layers:
            add_crypto_layers(psApp, main_doc, crypto_file)

    for layer in main_doc.Layers:
        if layer.Name in ("RGB", "A"):
            base_passes["layers"].append(layer)
        elif layer.Name.startswith('RGBA'):
            lighting_passes["layers"].append(layer)
        else:
            for keyword in render_pass_keywords:
                if keyword in layer.Name:
                    render_passes["layers"].append(layer)
                    break
            else:
                aovs_passes["layers"].append(layer)

    for group in (base_passes, render_passes, aovs_passes, lighting_passes):
        if len(group["layers"]) > 0:
            group_set = main_doc.LayerSets.Add()
            group_set.Name = group["name"]
            for layer in group["layers"]:
                layer.Move(group_set, 1)
    
    change_bit_detph(psApp, bit_value)
    psb = os.path.basename(exr_files[beauty_index]["path"])
    psb = psb.replace("_beauty", "")

    version = 0
    version_padding = f'{version:03}'
    psb = psb[:-4] + "_{}bit_{}.psb".format(bit_value, version_padding)

    while os.path.exists(psb):
        version += 1
        version_padding = f'{version:03}'
        psb = psb[:-7] + "{}.psb".format(version_padding)
        print(psb)
        
    psb = output + "\\" + psb
    save_psb(psApp, psb)
    print("Process Complete. {} saved.".format(psb))
    return psb