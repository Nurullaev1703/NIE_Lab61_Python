import xlsxwriter
guestnames = ['Игорь', 'Илья', 'Мади', 'Искандер']
try:
    my_file = ('visit.xlsx')
    for x in guestnames:
        book = xlsxwriter.Workbook(my_file)
        sheet = book.add_worksheet()
        sheet.set_column('B:B', 80)
        bold = book.add_format({'bold': True})
        sheet.write('B1', 'Новосибирский зоопарк', bold)
        sheet.write('B2', x + ', посетите')
        sheet.write('B3', 'один из лучших зоопарков стран СНГ',bold)
        sheet.write('B4', 'Только у нас вы сможете увидеть 750 видов животных, 350 из которых занесены в международную Красную книгу.')
        sheet.write('B5', 'Наш зоопарк также занимается разведением кошачьих и куньих, поэтому здесь одна из лучших в мире коллекций представителей этих семейств.')
        sheet.write('B6', 'Не ждите момента, а создавайте его!')
        sheet.write('B6', 'Мы ждем вас по адресу ул. Тимирязева, д.71/1', bold)
        sheet.insert_image('B7', 'zoo.jpg', {'x_scale': 0.50, 'y_scale': 0.50})
        print('Выполнено')
    book.close()

except Exception as a: # Обработка ошибок
    print("Error!")
    print(a)
