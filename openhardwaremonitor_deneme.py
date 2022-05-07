import wmi
from socket import timeout
import PySimpleGUI as sg  
import serial
import serial.tools.list_ports
w = wmi.WMI(namespace="root\OpenHardwareMonitor")
ser = serial.Serial()
sicaklik_verisi = w.Sensor()
sistem_verisi = w.Hardware()
sistem_dict = {}
sicaklik_dict = {}
load_dict = {}
power_dict = {}
data_dict = {}
#! Sayaçla ilgili değişkenler
sayac = 0
sayac_carpim = 10
bekleme_timeout = 1
toplam_sayfa = 3
for x in sistem_verisi:
    sistem_dict[x.HardwareType] = x.Name

#sistem_dict # Keys'HDD', 'CPU', 'GpuAti', 'Mainboard', 'RAM'
#sicaklik_dict  # Keys 'CPU Package', 'Temperature', 'CPU CCD #1'
#oad_dict #Keys 'CPU Core #1' ... 'CPU Core #6' 'CPU Total' 'Memory' 'Used Space' 
#power_dict #Keys 'CPU Core #1' ... 'CPU Core #6' 'CPU Cores' 'CPU Package' 'GPU Total'
#data_dict #Keys 'Used Memory' 'Available Memory' 
sg.theme("Reddit")

def serial_ports():
    ports = serial.tools.list_ports.comports()
    print(ports)
    seri_port = []
    for p in ports:
        print(p.device)
        seri_port.append(p.device)
    print(seri_port)
    return seri_port

def serial_baglan():
    com_deger = value[0]
    baud_deger = value[1]
    print("Baud Deger", value[1])
    global ser
    ser = serial.Serial(com_deger, baud_deger, timeout=0, parity=serial.PARITY_NONE, stopbits = serial.STOPBITS_ONE , bytesize = serial.EIGHTBITS, rtscts=0)
    window["-BAGLANDI_TEXT-"].update('Bağlandı...')

def nextion_metin_yazdir(metin, satir):
    eof =  b'\xff\xff\xff'
    metin = "t" + str(satir) + ".txt=\" " + metin + " \""
    ser.write(metin.encode('Ascii'))
    ser.write(eof)


baglanti_layout =[ [sg.Text("Port Seçiniz:"), sg.Combo(serial_ports(), size=(10,1)),
            sg.Text("Baud Seçiniz:"), sg.Combo(["110","300","600","1200", "2400", "4800", "9600", "14400", "19200", "38400", "57600", "115200", "128000", "256000"], default_value=9600), 
            sg.Button(button_text="Bağlan", key="-BAGLAN-", size=(10,1)) ],
            [sg.Text("", size=(10,1), key="-BAGLANDI_TEXT-")]
            ]
        
sistem_layout = [ #'HDD', 'CPU', 'GpuAti', 'Mainboard', 'RAM'
                 [sg.Text("İşlemci Adı:"), sg.Text("", key="islemci_adi")],
                 [sg.Text("GPU Adı:"), sg.Text("", key="gpu_adi")],
                 [sg.Text("Anakart:"), sg.Text("", key="anakart_adi")],
                 [sg.Text("RAM:"), sg.Text("", key="ram_adi")],
                 [sg.Text("Sabit Disk:"), sg.Text("", key="disk_adi")]
                ]
sicaklik_layout = [[sg.Text("CPU Package:"), sg.Text("", key="cpu_sicaklik")],
                   [sg.Text("Temperature:"), sg.Text("", key="temp_sicaklik")],
                   [sg.Text("CPU CCD #1:"), sg.Text("", key="ccd_sicaklik")]
                ]

load_layout = [ [sg.Text("CPU Core #1:"), sg.Text("", key="cpu_load")],
                [sg.Text("CPU Core #2:"), sg.Text("", key="cpu_load2")],
                [sg.Text("CPU Core #3:"), sg.Text("", key="cpu_load3")],
                [sg.Text("CPU Core #4:"), sg.Text("", key="cpu_load4")],
                [sg.Text("CPU Core #5:"), sg.Text("", key="cpu_load5")],
                [sg.Text("CPU Core #6:"), sg.Text("", key="cpu_load6")],
                [sg.Text("CPU Total:"), sg.Text("", key="cpu_total")],
                [sg.Text("Memory:"), sg.Text("", key="memory")],
                [sg.Text("Used Space:"), sg.Text("", key="used_space")]
                ]
power_layout = [[sg.Text("CPU Core #1:"), sg.Text("", key="cpu_power")],
                [sg.Text("CPU Core #2:"), sg.Text("", key="cpu_power2")],
                [sg.Text("CPU Core #3:"), sg.Text("", key="cpu_power3")],
                [sg.Text("CPU Core #4:"), sg.Text("", key="cpu_power4")],
                [sg.Text("CPU Core #5:"), sg.Text("", key="cpu_power5")],
                [sg.Text("CPU Core #6:"), sg.Text("", key="cpu_power6")],
                [sg.Text("CPU Package:"), sg.Text("", key="cpu_package")],
                [sg.Text("GPU Total:"), sg.Text("", key="gpu_total")]
                ]
data_layout = [ [sg.Text("Toplam RAM:"), sg.Text("", key="toplam_ram")],
                [sg.Text("Kullanılan RAM:"), sg.Text("", key="used_memory")],
                [sg.Text("Boş RAM:"), sg.Text("", key="available_memory")]
                ]


layout = [[sg.Frame("Bağlantı", baglanti_layout)],[sg.Frame("Sistem Bilgileri", sistem_layout) ,sg.Frame("Sıcaklık Bilgileri", sicaklik_layout)],
          [sg.Frame("Yük Bilgileri", load_layout) ,sg.Frame("Güç Bilgileri", power_layout)],
            [sg.Frame("RAM Bilgileri", data_layout)]]

window = sg.Window("Python Sistem Monitoru v2", layout)      



#!Ana program akışı
while True:                           
    event, value = window.read(timeout=bekleme_timeout) #100ms timeout

    # Olaylar 
    if event == sg.WIN_CLOSED:
        break      
    # * Program Döngüsü
    sicaklik_bilgisi = w.Sensor()
    for sensor in sicaklik_bilgisi:
        if sensor.SensorType == 'Temperature':
            sicaklik_dict[sensor.Name] = "{0:.2f}C".format(sensor.Value)
        elif sensor.SensorType == 'Load':
            load_dict[sensor.Name] = "%{0:.2f}".format(sensor.Value)
        elif sensor.SensorType == 'Power':
            power_dict[sensor.Name] = "{0:.2f}W".format(sensor.Value)
        elif sensor.SensorType == 'Data':
            data_dict[sensor.Name] = "{0:05.2f}GB".format(sensor.Value)
    if event == "-BAGLAN-":
        if (value[0] == ""):
            sg.popup("Bir Port Seçiniz!", title="Hata", custom_text="Tamam") 
        elif (value[1] == ""):
            sg.popup("Baud Oranını Seçiniz!", title="Hata", custom_text="Tamam")
        else:
            serial_baglan()
    # Bilgileri Ekrana Yazdırma 
    window["islemci_adi"].update(sistem_dict["CPU"])
    window["gpu_adi"].update(sistem_dict["GpuAti"])
    window["anakart_adi"].update(sistem_dict["Mainboard"])
    window["ram_adi"].update(sistem_dict["RAM"])
    window["disk_adi"].update(sistem_dict["HDD"])
    window["cpu_sicaklik"].update(sicaklik_dict["CPU Package"])
    window["temp_sicaklik"].update(sicaklik_dict["Temperature"])
    window["ccd_sicaklik"].update(sicaklik_dict["CPU CCD #1"])
    window["cpu_load"].update(load_dict["CPU Core #1"])
    window["cpu_load2"].update(load_dict["CPU Core #2"])
    window["cpu_load3"].update(load_dict["CPU Core #3"])
    window["cpu_load4"].update(load_dict["CPU Core #4"])
    window["cpu_load5"].update(load_dict["CPU Core #5"])
    window["cpu_load6"].update(load_dict["CPU Core #6"])
    window["cpu_total"].update(load_dict["CPU Total"])
    window["memory"].update(load_dict["Memory"])
    window["used_space"].update(load_dict["Used Space"])
    window["cpu_power"].update(power_dict["CPU Core #1"])
    window["cpu_power2"].update(power_dict["CPU Core #2"])
    window["cpu_power3"].update(power_dict["CPU Core #3"])
    window["cpu_power4"].update(power_dict["CPU Core #4"])
    window["cpu_power5"].update(power_dict["CPU Core #5"])
    window["cpu_power6"].update(power_dict["CPU Core #6"])
    window["cpu_package"].update(power_dict["CPU Package"])
    window["gpu_total"].update(power_dict["GPU Total"])
    window["toplam_ram"].update("{0:05.2f}GB".format(float(data_dict["Used Memory"][:-2]) + float(data_dict["Available Memory"][:-2])))
    window["used_memory"].update(data_dict["Used Memory"])
    window["available_memory"].update(data_dict["Available Memory"])

    # Nextion Ekrana Yazdırma
    if ser.isOpen():
        if sayac < sayac_carpim * 1:
            nextion_metin_yazdir("SISTEM BILGISI", 0)
            nextion_metin_yazdir("Islemci:" + sistem_dict["CPU"], 1)
            nextion_metin_yazdir("GPU:" + sistem_dict["GpuAti"], 2)
            nextion_metin_yazdir("Anakart:" + sistem_dict["Mainboard"], 3)
            nextion_metin_yazdir("RAM:" + sistem_dict["RAM"], 4)
            nextion_metin_yazdir("Disk:" + sistem_dict["HDD"], 5)
            nextion_metin_yazdir("RAM Boyut:" + "{0:05.2f}GB".format(float(data_dict["Used Memory"][:-2]) + float(data_dict["Available Memory"][:-2])), 6)
            nextion_metin_yazdir("SICAKLIK VERISI", 7)
            nextion_metin_yazdir("CPU Sicaklik:" + sicaklik_dict["CPU Package"], 8)
            nextion_metin_yazdir("Temp:" + sicaklik_dict["Temperature"], 9)
            nextion_metin_yazdir("CCD:" + sicaklik_dict["CPU CCD #1"], 10)
        elif sayac < sayac_carpim * 2:
            nextion_metin_yazdir("YUK VERISI", 0)
            nextion_metin_yazdir("CPU Core 1:" + load_dict["CPU Core #1"], 1)
            nextion_metin_yazdir("CPU Core 2:" + load_dict["CPU Core #2"], 2)
            nextion_metin_yazdir("CPU Core 3:" + load_dict["CPU Core #3"], 3)
            nextion_metin_yazdir("CPU Core 4:" + load_dict["CPU Core #4"], 4)
            nextion_metin_yazdir("CPU Core 5:" + load_dict["CPU Core #5"], 5)
            nextion_metin_yazdir("CPU Core 6:" + load_dict["CPU Core #6"], 6)
            nextion_metin_yazdir("CPU Core 7:" + load_dict["CPU Total"], 7)
            nextion_metin_yazdir("RAM Yuk:" + load_dict["Memory"], 8)
            nextion_metin_yazdir("Disk Yuk:" + load_dict["Used Space"], 9)
            nextion_metin_yazdir("", 10)
        elif sayac < sayac_carpim * 3:
            nextion_metin_yazdir("GUC VERISI", 0)
            nextion_metin_yazdir("CPU Core 1:" + power_dict["CPU Core #1"], 1)
            nextion_metin_yazdir("CPU Core 2:" + power_dict["CPU Core #2"], 2)
            nextion_metin_yazdir("CPU Core 3:" + power_dict["CPU Core #3"], 3)
            nextion_metin_yazdir("CPU Core 4:" + power_dict["CPU Core #4"], 4)
            nextion_metin_yazdir("CPU Core 5:" + power_dict["CPU Core #5"], 5)
            nextion_metin_yazdir("CPU Core 6:" + power_dict["CPU Core #6"], 6)
            nextion_metin_yazdir("CPU Total:" + power_dict["CPU Package"], 7)
            nextion_metin_yazdir("GPU:" + power_dict["GPU Total"], 8)
            nextion_metin_yazdir("", 9)
            nextion_metin_yazdir("", 10)
        sayac += 1
        if(sayac > toplam_sayfa * sayac_carpim):
            sayac = 0

window.close()
