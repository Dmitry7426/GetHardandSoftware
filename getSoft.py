import subprocess
import re
import time
from win32com.client import Dispatch

while True:
    act = input('Введите имя ПК (IP адрес), с которого необходимо собрать информацию: \n '
                'Если поле оставить пустым - информация будет собрана на локальном ПК \b -> ')
    if len(act) == 0:
        server = Dispatch("WbemScripting.SWbemLocator")
        c = server.ConnectServer(act, "root\\cimv2")
        comp_system = c.ExecQuery(
            "Select Name, UserName from Win32_ComputerSystem")  # Получаем имя ПК и текущего пользователя
        for i in comp_system[0].Properties_:
            if i.Name == 'Name':
                act = i.Value
    print('Информация будет собрана для ПК', act)
    slovar = list()
    server = Dispatch("WbemScripting.SWbemLocator")
    try:
        c = server.ConnectServer(act, "root\\cimv2")
        comp_system = c.ExecQuery("Select Name, "
                                  "UserName from Win32_ComputerSystem")  # Получаем имя ПК и текущего пользователя
        comp_network_speed = c.ExecQuery("Select NetConnectionStatus, Speed, "
                                         "MACAddress from Win32_NetworkAdapter")  # Получаем информацию о сети
        comp_network_conf = c.execQuery("Select IPAddress, Description, MACAddress "
                                        "from Win32_NetworkAdapterConfiguration")  # Получаем информацию о сети
        comp_os = c.execQuery("Select Caption, TotalVisibleMemorySize "
                              "from Win32_OperatingSystem")  # Получаем информацию о ОС
        comp_board = c.execQuery("Select Manufacturer, Product "
                                 "from Win32_BaseBoard")  # Получаем информацию о системной плате
        comp_processor = c.execQuery("Select Name, SocketDesignation "
                                     "from Win32_Processor")  # Получаем информацию о процессоре
        comp_hdd = c.execQuery("Select Caption, Size "
                               "from Win32_DiskDrive")  # Получаем информацию о дисках
        comp_video = c.execQuery("Select Description from Win32_videoController")  # Получаем информацию о видео
        comp_printer = c.execQuery("Select Caption, PortName from Win32_Printer")  # Получаем информацию о принтерах
        comp_software = c.execQuery("Select Name from Win32_Product")  # Получаем информацию о принтерах

        class get_sysinfo:

            @staticmethod
            def name_pc():  # Получаем имя ПК и текущего пользователя
                for q in comp_system[0].Properties_:
                    if q.Name == 'Name':
                        slovar.append({'Имя компьютера': q.Value})
                    if 'UserName' in q.Name:
                        slovar.append({'Имя текущего пользователя': q.Value})

            @staticmethod
            def get_network():
                e = 0
                for _ in comp_network_speed:
                    for q in comp_network_speed[e].Properties_:
                        # print(i.Name, i.Value)
                        if 'Speed' in q.Name and q.Value:
                            slovar.append({f'{q.Name}': int(q.Value) / 1000 / 1000})
                    e += 1
                e = 0
                for _ in comp_network_conf:
                    for q in comp_network_conf[e].Properties_:
                        if 'IPAddress' in q.Name and q.Value:
                            slovar.append({f'{q.Name}': q.Value})
                    e += 1

            @staticmethod
            def get_os():
                for q in comp_os[0].Properties_:
                    if 'TotalVisibleMemorySize' in q.Name:
                        mem = round(int(q.Value) / 1024 / 1024)
                        slovar.append({'ОЗУ': f'{mem} Gb'})
                    if 'Caption' in q.Name:
                        slovar.append({'Установленная ОС': q.Value})

            @staticmethod
            def get_board():
                for q in comp_board[0].Properties_:
                    slovar.append({f'{q.Name}': q.Value})

            @staticmethod
            def get_processor():
                for q in comp_processor[0].Properties_:
                    slovar.append({f'{q.Name}': q.Value})

            @staticmethod
            def get_hdd():
                e = 0
                for _ in comp_hdd:
                    for q in comp_hdd[e].Properties_:
                        if 'Caption' in q.Name:
                            slovar.append({'Модель HDD': q.Value})
                        if 'Size' in q.Name:
                            slovar.append({'Объем HDD': f'{round(int(q.Value) / 1024 / 1024 / 1024)} Gb'})
                    e += 1

            @staticmethod
            def get_video():
                e = 0
                for _ in comp_video:
                    for q in comp_video[e].Properties_:
                        slovar.append({f'{q.Name}': q.Value})
                    e += 1

            @staticmethod
            def get_printer():
                e = 0
                for _ in comp_printer:
                    for q in comp_printer[e].Properties_:
                        if 'Caption' in q.Name:
                            slovar.append({'Наименование принтера': q.Value})
                        if 'PortName' in q.Name:
                            slovar.append({'Порт принтера': q.Value})
                    e += 1

            @staticmethod
            def get_soft():
                e = 0
                for _ in comp_software:
                    for q in comp_software[e].Properties_:
                        if q.Name == 'Name' and q.Name != 'Имя':
                            slovar.append({f'{q.Name}': q.Value})
                    e += 1

            @staticmethod
            def get_activate():
                activates = subprocess.Popen("powershell.exe -ExecutionPolicy ByPass "
                                             "-File act.ps1", stdout=subprocess.PIPE)
                w = str(activates.communicate())
                w = re.findall(r'\d+', w)
                if '1' in w:
                    activation = 'Система активирована'
                else:
                    activation = 'Система не активирована'
                slovar.append({'Статус активации': activation})

            @staticmethod
            def file_output():
                with open(f'//winwsus/inv21/555/{act}.txt', 'a+',
                          encoding='UTF-8') as file_out:
                    for q in slovar:
                        for key, val in q.items():
                            file_out.writelines(f'{key}: {val}; \n')

        get_sysinfo.name_pc()
        get_sysinfo.get_os()
        get_sysinfo.get_activate()
        get_sysinfo.get_network()
        get_sysinfo.get_processor()
        get_sysinfo.get_board()
        get_sysinfo.get_hdd()
        get_sysinfo.get_video()
        get_sysinfo.get_printer()
        get_sysinfo.get_soft()
        get_sysinfo.file_output()
        print(f'Информация о {act} успешно собрана. Данные сохранены в БД.')
        time.sleep(10)

        break

    except Exception as ers:
        print('\nИмя ПК не доступно, либо имя введено не верно. Попробуйте еще раз! \n', ers)
