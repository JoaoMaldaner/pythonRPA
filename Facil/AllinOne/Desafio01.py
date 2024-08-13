from selenium.webdriver.common.by import By
import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

class Desafio01:

    def __init__(self) -> None:
        self.nav = webdriver.Chrome()
        self.nav.get('https://webscraper.io/test-sites/e-commerce/static')
        self.nav.maximize_window()
    def element_save(self, element_data):
        elements = self.nav.find_elements(By.CLASS_NAME, "thumbnail")
        for produto in elements:
            nome = produto.find_element(By.XPATH, ".//a[@title]").get_attribute("title").strip()
            preco = produto.find_element(By.CLASS_NAME, "price").text
            descricao = produto.find_element(By.CLASS_NAME, "description").text
            estrelas = produto.find_element(By.XPATH, ".//div[@class='ratings']//p[@data-rating]").get_attribute(
                "data-rating")
            reviews = produto.find_element(By.XPATH, ".//p[@class='review-count float-end']").text.strip()

            element_data.append({
                "Nome": nome,
                "Preco": preco,
                "Descricao": descricao,
                "Estrelas": estrelas,
                "Reviews": reviews
            })
        return element_data
    def columns(self, df):
        # Converte as colunas para os dados apropriados
        df['Preco'] = df['Preco'].replace(r'[\$,]', '', regex=True).astype(float)
        df['Reviews'] = df['Reviews'].replace(' reviews', '', regex=True).astype(int)
        df['Estrelas'] = df['Estrelas'].astype(int)
        return df
    def ordElements(self, df):

        valor = df.sort_values(by='Preco', ascending=True)
        qtdReviews = df.sort_values(by='Reviews', ascending=False)
        qtdEstrelas = df.sort_values(by='Estrelas', ascending=False)
        bonus = df.sort_values(by=['Preco', 'Reviews', 'Estrelas'], ascending=[True, False, False])
        return valor, qtdReviews, qtdEstrelas, bonus
    def excel(self, produtos_ordenados, nome_arquivo='produtos.xlsx'):
        with pd.ExcelWriter(nome_arquivo) as writer:
            produtos_ordenados[0].to_excel(writer, sheet_name='Preço', index=False)
            produtos_ordenados[1].to_excel(writer, sheet_name='Reviews', index=False)
            produtos_ordenados[2].to_excel(writer, sheet_name='Estrelas', index=False)
            produtos_ordenados[3].to_excel(writer, sheet_name='Bonus', index=False)

    def iniciar(self) -> None:

        cookies = self.nav.find_element(By.XPATH, '//*[@id="cookieBanner"]/div[2]/a').click()

        WebDriverWait(self.nav, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "category-link")))
        computers = self.nav.find_element(By.XPATH, "//a[@href='/test-sites/e-commerce/static/computers']").click()

        clickLaptops = self.nav.find_element(By.XPATH, "//a[@href='/test-sites/e-commerce/static/computers/laptops']").click()
        WebDriverWait(self.nav, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "thumbnail")))
        laptopSave = []
        self.element_save(laptopSave)
        df_laptops = pd.DataFrame(laptopSave)
        df_laptops = self.columns(df_laptops)
        laptopsOrd = self.ordElements(df_laptops)
        self.excel(laptopsOrd, nome_arquivo='laptops_desafio01.xlsx')

        clickTablets = self.nav.find_element(By.XPATH, "//a[@href='/test-sites/e-commerce/static/computers/tablets']").click()
        WebDriverWait(self.nav, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "thumbnail")))
        tabletSave = []
        self.element_save(tabletSave)
        df_tablets = pd.DataFrame(tabletSave)
        df_tablets = self.columns(df_tablets)
        tabletsOrd = self.ordElements(df_tablets)
        self.excel(tabletsOrd, nome_arquivo='tablets_desafio02.xlsx')

        phones = self.nav.find_element(By.XPATH, "//a[@href='/test-sites/e-commerce/static/phones']").click()

        clickTouch = self.nav.find_element(By.XPATH, "//a[@href='/test-sites/e-commerce/static/phones/touch']").click()

        WebDriverWait(self.nav, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "thumbnail")))
        phoneSave = []
        self.element_save(phoneSave)
        df_phones = pd.DataFrame(phoneSave)
        df_phones = self.columns(df_phones)
        phonesOrd = self.ordElements(df_phones)
        self.excel(phonesOrd, nome_arquivo='phones_desafio01.xlsx')
        input()


if __name__ == '__main__':
    Desafio01().iniciar()