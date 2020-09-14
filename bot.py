from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from editar import user, password, initialDate, finalDate
from os import listdir
import pandas as pd


class bot():
    def __init__(self):
        self.driver = webdriver.Chrome()

    def login(self):
        print("---------------------------------------------")
        print("Entrando a helios")
        # abre la pagina
        self.driver.get('http://helios.apolloecommerce.com/')

        sleep(2)

        user_in = self.driver.find_element_by_xpath('//*[@id="root"]/div[2]/div/div/div/div[2]/form/div[1]/input')  # selecciono el input de usuario
        user_in.send_keys(user)  # Envio los datos de usuario

        pw_in = self.driver.find_element_by_xpath('//*[@id="root"]/div[2]/div/div/div/div[2]/form/div[2]/input')  # selecciono el input de password
        pw_in.send_keys(password)  # Envio los datos de password

        log_btn = self.driver.find_element_by_xpath('//*[@id="root"]/div[2]/div/div/div/div[2]/form/div[4]/button')  # selecciono el boto de login
        log_btn.click()  # doy click

        sleep(5)


    def akaesColombia(self): 

        print("---------------------------------------------")
        print("Logeando a AkaesColombia")

        sel_cuenta_btn= self.driver.find_element_by_xpath('//*[@id="root"]/nav/div/ul/div[1]/button') 
        sel_cuenta_btn.click()

        akaes_btn=bot.driver.find_element_by_xpath('//button[normalize-space()="AKAESCOLOMBIA"]')
        akaes_btn.click()

        sleep(1)
       
        reportes_a=self.driver.find_element_by_xpath('//*[@id="root"]/nav/div[1]/ul/div[3]/a')
        reportes_a.click()

        ordenes_btn=self.driver.find_element_by_xpath('//*[@id="root"]/nav/div[1]/ul/div[3]/div/button[2]')
        ordenes_btn.click()

        sleep(2)

        fecha_inicio=self.driver.find_element_by_xpath('//*[@id="root"]/div[2]/div/div[2]/form/div[1]/div/div/input')
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(Keys.BACKSPACE)
        fecha_inicio.send_keys(initialDate)

        fecha_fin=self.driver.find_element_by_xpath('//*[@id="root"]/div[2]/div/div[2]/form/div[2]/div/div/input')
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(Keys.BACKSPACE)
        fecha_fin.send_keys(finalDate)

        generar_btn=self.driver.find_element_by_xpath('//*[@id="root"]/div[2]/div/div[2]/form/div[3]/button')
        generar_btn.click()

bot=bot()
bot.login()
bot.akaesColombia()
