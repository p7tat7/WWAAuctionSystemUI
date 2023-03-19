import wwaglobal
import wwalog
import wwautils
import wwafindbidder
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import Calendar
from datetime import datetime
import time
import tkcap
import test
import wwaformatcreate
wwaglobal.init()


# LARGEFONT = ("Verdana", 35)
def check_order_list():

    wwautils.checkOrderList()
    time.sleep(1)
def exit():
    app.withdraw()

class tkinterApp(tk.Tk):

    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        self.geometry('300x300')
        self.title('Auction System')
        # creating a container
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # initializing frames to an empty array
        self.frames = {}

        # iterating through a tuple consisting
        # of the different page layouts
        for F in (StartPage, Auction, find_bidder, end_auction, cms_order,CMS, whatsapp, post_auction,message_notification):
            frame = F(container, self)

            # initializing frame of that object from
            # startpage, page1, page2 respectively with
            # for loop
            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartPage)

    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


# first window frame startpage

class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        # label of frame Layout 2
        label = ttk.Label(self, text="Auction System")

        # putting the grid in its place by using
        # grid
        label.grid(row=0, column=1, padx=10, pady=10)

        button1 = ttk.Button(self, text="拍賣系統",
                             command=lambda: controller.show_frame(Auction))

        # putting the button in its place by
        # using grid

        button1.grid(row=1, column=1, padx=5, pady=5)

        ## button to show frame 2 with text layout2
        button2 = ttk.Button(self, text="CMS 網店查詢",
                             command=lambda: controller.show_frame(CMS))

        # putting the button in its place by
        # using grid
        button2.grid(row=2, column=1, padx=5, pady=5)
        button3 = ttk.Button(self, text='Whatsapp 功能', command=lambda: controller.show_frame(whatsapp))
        button3.grid(row=3, column=1, padx=5, pady=5)
        button4 = ttk.Button(self, text='查看交收物品位置', command=lambda: wwautils.check_auction_item())
        button4.grid(row=4, column=1, padx=5, pady=5)
        button5 = ttk.Button(self, text='Exit', command=lambda: self.quit())
        button5.grid(row=5, column=1, padx=5, pady=5)
        button6 = ttk.Button(self, text='分配位置', command=lambda: test.auction_place())
        button6.grid(row=6, column=1, padx=5, pady=5)
        button7 = ttk.Button(self, text='TEST', command=lambda: test.test_post_auction())
        button7.grid(row=9, column=1, padx=5, pady=5)

# second window frame page1
class Auction(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = ttk.Label(self, text="拍賣")
        label.grid(row=0, column=1, padx=5, pady=5)

        # button to show frame 2 with text
        # layout2
        button1 = ttk.Button(self, text="Back",
                             command=lambda: controller.show_frame(StartPage))

        # putting the button in its place
        # by using grid
        button1.grid(row=1, column=1, padx=5, pady=5)

        # button to show frame 2 with text
        # layout2
        button2 = ttk.Button(self, text="出拍賣",
                             command=lambda: controller.show_frame(post_auction))

        # putting the button in its place by
        # using grid
        button2.grid(row=2, column=1, padx=10, pady=10)
        button3 = ttk.Button(self, text='記錄中標者', command=lambda: controller.show_frame(find_bidder))
        button3.grid(row=3, column=1, padx=5, pady=5)
        button4 = ttk.Button(self, text='完標', command= lambda:controller.show_frame(end_auction))
        button4.grid(row=4, column=1, padx=5, pady=5)
        button5 = ttk.Button(self, text='催拍賣中標者落單', command=lambda: wwautils.urgeOrder())
        button5.grid(row=5, column=1, padx=5, pady=5)
        button6 = ttk.Button(self, text='排列拍賣品', command=lambda :controller.show_frame(cms_order))
        button6.grid(row=6, column=1, padx=5, pady=5)

# third window frame page2
class post_auction(tk.Frame):

    def __init__(self, parent, controller):
        def post():
            postnumber = int(postnumber1.get())
            postnumber_fever = int(postnumber2.get())
            wwautils.postAuction(postnumber, postnumber_fever)
            controller.show_frame(Auction)
        tk.Frame.__init__(self, parent)
        button1 = tk.Button(self, text='Back', command=lambda:controller.show_frame(Auction))
        button1.grid(row=0,column=1)
        label1 = tk.Label(self, text='環球水族拍賣區數量:')
        label1.grid(row=1, column=1)
        postnumber1 = tk.Entry(self)
        postnumber1.grid(row=1,column=2,pady=5)
        label2 = tk.Label(self, text='發燒友拍賣區:')
        label2.grid(row=2, column=1)
        postnumber2 = tk.Entry(self)
        postnumber2.grid(row=2,column=2,pady=5)
        button2 = tk.Button(self, text='Post', command=lambda:post())
        button2.grid(row=3,column=2)

class find_bidder(tk.Frame):

    def __init__(self, parent, controller):
        def get_date():
            date = cal.get_date()
            #print(datetime.strptime(date,'%d/%m/%Y'))

            nameList = wwafindbidder.findbidder(date)
            f = open(wwaglobal.src_path + "nameList.txt", "w+", encoding = "utf-8")
            for i in range(len(nameList)):
                if not nameList[i] == "":
                    f.write(nameList[i] + "\n")
            f.close()
            controller.show_frame(Auction)
        tk.Frame.__init__(self, parent)

        label = ttk.Label(self, text="請選擇需完標的拍賣發佈日期")
        label.grid(row=1, column=1, padx=10, pady=10)
        cal = Calendar(self, selectmode='day',
                       year=int(datetime.today().year), month=int(datetime.today().month),
                       day=int(datetime.today().day),date_pattern='y-mm-dd')
        #date_entry = tk.Entry(self)
        cal.grid(row=2, column=1, padx=5,pady=5)
        # button to show frame 2 with text
        # layout2
        button1 = ttk.Button(self, text="Back",
                             command=lambda: controller.show_frame(Auction))

        # putting the button in its place by
        # using grid
        button1.grid(row=0, column=1, padx=10, pady=10)
        button2 = tk.Button(self, text='查找中標者', command=lambda:get_date())
        button2.grid(row=3, column=1)



class end_auction(tk.Frame):
    def __init__(self, parent, controller):
        def end():
            date = cal.get_date()
            wwautils.endAuction(date)
            controller.show_frame(Auction)
        tk.Frame.__init__(self, parent)
        button1 = ttk.Button(self, text='Back', command=lambda:controller.show_frame(Auction))
        button1.grid(row=0, column=1, padx=5, pady=5)
        label = ttk.Label(self, text="請選擇需完標的拍賣發佈日期")
        label.grid(row=1, column=1, padx=10, pady=10)
        cal = Calendar(self, selectmode='day',
                       year=int(datetime.today().year), month=int(datetime.today().month),
                       day=int(datetime.today().day), date_pattern='y-mm-dd')
        # date_entry = tk.Entry(self)
        cal.grid(row=2, column=1, padx=5, pady=5)
        button2 = tk.Button(self, text='執行', command=lambda: end())
        button2.grid(row=3, column=1)

class cms_order(tk.Frame):
    def __init__(self, parent, controller):
        def run():
            try:
                number = int(entry1.get())

                wwautils.move_objects_cms(number)
                controller.show_frame(Auction)
            except ValueError:
                error_box('Input Error')

        tk.Frame.__init__(self, parent)
        button1 = ttk.Button(self, text='Back', command=lambda: controller.show_frame(Auction))
        button1.grid(row=1, column=1, padx=5, pady=5)
        label1 = tk.Label(self, text='需要置底多少商品: ')
        label1.grid(row=2, column=1, padx=5, pady=10)
        entry1 = tk.Entry(self)
        entry1.grid(row=2, column=2, padx=5, pady=10)
        button2 = tk.Button(self, text='Run', command=lambda: run())
        button2.grid(row=3, columnspan=2, column=1, padx=5, pady=10)
class CMS(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        button1 = ttk.Button(self, text='Back', command=lambda: controller.show_frame(StartPage))
        button1.grid(row=1, column=1, padx=5, pady=5)
        button2 = ttk.Button(self, text='入交收list', command=lambda:check_order_list())
        button2.grid(row=2, column=1, padx=5, pady=5)

class whatsapp(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)
        button1 = ttk.Button(self, text='Back', command=lambda: controller.show_frame(StartPage))
        button1.grid(row=1, column=1, padx=5, pady=5)
        button2 = ttk.Button(self, text='提醒今日交收', command=lambda: controller.show_frame(message_notification))
        button2.grid(row=2, column=1, padx=5, pady=5)
        button3 = ttk.Button(self, text='售後慰問', command=lambda: wwautils.PostOrder())
        button3.grid(row=3, column=1, padx=5, pady=5)
        button4 = ttk.Button(self, text='完成交收', command=lambda: wwautils.finished_meeting())
        button4.grid(row=4, column=1, padx=5, pady=5)

class message_notification(tk.Frame):
    def __init__(self, parent, controller):
        def message(staff):
            wwautils.meetingNotification(staff)
            controller.show_frame(whatsapp)
        tk.Frame.__init__(self, parent)
        button1 = ttk.Button(self, text='Back', command=lambda: controller.show_frame(whatsapp))
        button1.grid(row=1, column=1, padx=5, pady=5)
        label1 = tk.Label(self, text='身份')
        label1.grid(row=2, column=1, padx=15, pady=5)
        button2 = tk.Button(self, text='George', command=lambda: message(1))
        button2.grid(row=3, column=1, padx=15, pady=5)
        button3 = tk.Button(self, text='Harry', command=lambda: message(2))
        button3.grid(row=4, column=1, padx=15, pady=5)
        button4 = tk.Button(self, text='Ryan', command=lambda: message(3))
        button4.grid(row=5, column=1, padx=15, pady=5)
# Driver Code

def phone_window(phone):
    def exited():
        list_phone.destroy()
        list_phone.quit()
    def disable_event():
        pass
    list_phone = tk.Toplevel()
    list_phone.title('Phone Numbers')
    list_phone.geometry('400x200')
    label1 = tk.Label(list_phone, text='請確保以下電話號碼已加入聯絡人 確認後請按Close以繼續', anchor='center')
    label1.pack()
    list = ttk.Treeview(list_phone, column=('c1', 'c2'), show='headings', height=5)
    list.column('# 1', anchor='center')
    list.heading('# 1', text='Phone')
    list.column('# 2', anchor='center')
    list.heading('# 2', text='Name')
    # list.pack()
    # list['columns'] = ('Phone', 'Name')
    # +list.column('#0',text='',anchor='center')
    # list.heading("Phone", text="Phone", anchor='center')
    # list.heading("Name", text="Name", anchor='center')
    for index, phone_number in enumerate(phone):
        list.insert('', 'end', text=index, values=(phone_number[0], phone_number[1]))
    list.pack()
    #exited = False
    button1 = tk.Button(list_phone, text='Close', command = lambda: exited(), anchor='center')
    button1.pack()
    list_phone.protocol("WM_DELETE_WINDOW", disable_event)

    list_phone.mainloop()

def order_id_window(order):
    def exited():
        list_order.destroy()
        list_order.quit()
    def disable_event():
        pass
    list_order = tk.Toplevel()
    list_order.title('Order ID')
    list_order.geometry('400x200')
    label1 = tk.Label(list_order, text='狀態異常或查找不到的訂單', anchor='center', pady=5)
    label1.pack()
    list = ttk.Treeview(list_order, column=('c1', 'c2'), show='headings', height=5)
    list.column('# 1', anchor='center')
    list.heading('# 1', text='訂單編號')
    list.column('# 2', anchor='center')
    list.heading('# 2', text='錯誤訊息')
    # list.pack()
    # list['columns'] = ('Phone', 'Name')
    # +list.column('#0',text='',anchor='center')
    # list.heading("Phone", text="Phone", anchor='center')
    # list.heading("Name", text="Name", anchor='center')
    for index, invalid_order in enumerate(order):
        if invalid_order[1] == 'invalid status':
            list.insert('', 'end', text=index, values=(str(invalid_order[0]), '狀態異常'))
        elif invalid_order[1] == 'not found':
            list.insert('', 'end', text=index, values=(str(invalid_order[0]), '無法查找'))
    list.pack()
    #exited = False
    button1 = tk.Button(list_order, text='Close', command = lambda: exited(), anchor='center')
    button1.pack()
    list_order.protocol("WM_DELETE_WINDOW", disable_event)

    list_order.mainloop()

def placement_window(auction_list, placement_list, area_b_list,remain):
    def save():
        try:
            cap = tkcap.CAP(placement)
            filename = f'{wwaglobal.src_placement}\\珊瑚分配位置\\{datetime.today().year}-{datetime.today().month}-{datetime.today().day} {datetime.today().hour}{datetime.today().minute}{datetime.today().second}.jpg'
            print(filename)
            cap.capture(filename)
            messagebox.showinfo('Complete', f'Saved successfully in {filename}')
        except:
            messagebox.showerror('Error', 'Unsuccessful to save.')
    def exited():
        placement.destroy()
        placement.quit()
    def disable_event():
        pass
    placement = tk.Toplevel()
    placement.title('Placement')
    placement.geometry('1700x750')
    label1 = tk.Label(placement, text='中標珊瑚', anchor='center')
    label1.grid(row=0, column=1, pady=5)
    label2 = tk.Label(placement, text='座標分配', anchor='center')
    label2.grid(row=0, column=2, pady=5)
    label3 = tk.Label(placement, text='以下珊瑚放至B區', anchor='center')
    label3.grid(row=0, column=3, pady=5)
    label1 = tk.Label(placement, text='未分配位置', anchor='center')
    label1.grid(row=0, column=4, pady=5)

    list1 = ttk.Treeview(placement, column=('c1'), show='headings', height=30)
    list1.column('# 1', anchor='center')
    list1.heading('# 1', text='中標珊瑚')
    for index, item in enumerate(auction_list):
        list1.insert('', 'end', text=index, values='CAA'+item[0])
    list1.grid(row=1, column=1,padx=5, pady=5)

    list2 = ttk.Treeview(placement, column=('c1', 'c2', 'c3'), show='headings', height=30)
    list2.column('# 1', anchor='center')
    list2.heading('# 1', text='座標')
    list2.column('# 2', anchor='center')
    list2.heading('# 2', text='珊瑚編號')
    list2.column('# 3', anchor='center')
    list2.heading('# 3', text='品種')
    for index, item in enumerate(placement_list):
        if len(item) != 3:
            break
        list2.insert('', 'end', text=index, values=(item[0], 'CAA'+item[1], item[2]))
    list2.grid(row=1, column=2, padx=5, pady=5)

    list3 = ttk.Treeview(placement, column=('c1', 'c2'), show='headings', height=30)
    list3.column('# 1', anchor='center')
    list3.heading('# 1', text='珊瑚編號')
    list3.column('# 2', anchor='center')
    list3.heading('# 2', text='品種')
    for index, item in enumerate(area_b_list):
        list3.insert('', 'end', text=index, values=('CAA'+item[0], item[1]))
    list3.grid(row=1, column=3, padx=5, pady=5)

    list4 = ttk.Treeview(placement, column=('c1', 'c2'), show='headings', height=30)
    list4.column('# 1', anchor='center')
    list4.heading('# 1', text='珊瑚編號')
    list4.column('# 2', anchor='center')
    list4.heading('# 2', text='品種')
    for index, item in enumerate(remain):
        list4.insert('', 'end', text=index, values=('CAA' + item[0], item[1]))
    list4.grid(row=1, column=4, padx=5, pady=5)

    button1 = tk.Button(placement, text='Save', command=lambda: save())
    button1.grid(row=2, column=3)
    button2 = tk.Button(placement, text='Close', command=lambda: exited(), anchor='center')
    button2.grid(row=2, column=2)
    placement.protocol("WM_DELETE_WINDOW", disable_event)
    placement.mainloop()


def check_placement_window(coordinates, unfound):
    def save():
        try:
            cap = tkcap.CAP(place_list)
            filename = f'{wwaglobal.src_placement}\\位置查找紀錄\\{datetime.today().year}-{datetime.today().month}-{datetime.today().day} {datetime.today().hour}{datetime.today().minute}{datetime.today().second}.jpg'
            print(filename)
            cap.capture(filename)
            messagebox.showinfo('Complete', f'Saved successfully in {filename}')
        except:
            messagebox.showerror('Error', 'Unsuccessful to save.')
    def exited():
        place_list.destroy()
        place_list.quit()
    def disable_event():
        pass
    place_list = tk.Toplevel()
    place_list.title('Check Placement')
    place_list.geometry('1000x750')
    label1 = tk.Label(place_list, text='珊瑚位置', anchor='center', pady=5)
    label1.grid(row=0, column=1)

    list1 = ttk.Treeview(place_list, column=('c1', 'c2','c3'), show='headings', height=30)
    list1.column('# 1', anchor='center')
    list1.heading('# 1', text='索引')
    list1.column('# 2', anchor='center')
    list1.heading('# 2', text='珊瑚編號')
    list1.column('# 3', anchor='center')
    list1.heading('# 3', text='座標')
    for index, content in enumerate(coordinates):
        list1.insert('', 'end', text=index, values=(content[0], content[1], content[2]))
    list1.grid(row=1,column=1)

    list2 = ttk.Treeview(place_list, column=('c1'), show='headings', height=30)
    list2.column('# 1', anchor='center')
    list2.heading('# 1', text='珊瑚編號')
    for index, content in enumerate(unfound):
        list2.insert('', 'end', text=index, values=(content))
    list2.grid(row=1, column=2)

    #exited = False
    button1 = tk.Button(place_list, text='Save', command = lambda: save(), anchor='center')
    button1.grid(row=2, column=1)
    button2 = tk.Button(place_list, text='Close', command=lambda: exited(), anchor='center')
    button2.grid(row=2, column=2)
    place_list.protocol("WM_DELETE_WINDOW", disable_event)
    place_list.mainloop()

def error_box(message):
    messagebox.showerror('Error', message)

def question_box(title, message):
    answer = tk.messagebox.askquestion(title, message)
    #print(answer)
    return answer
def main():
    app = tkinterApp()
    app.mainloop()
