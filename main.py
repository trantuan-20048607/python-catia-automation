import win32com.client

if __name__ == '__main__':
    catia = win32com.client.Dispatch('catia.application')
    doc = catia.activedocument
