import win32com.client as win32

# send mail from the account that is currently logged in

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'testmail@mail.com'
mail.Subject = 'Mail subject'
mail.Body = 'Mail body...' # for plain text body
#mail.HTMLBody = '' # for HTML body
mail.Display(True) # 
#mail.send