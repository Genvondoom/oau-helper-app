import os
import pandas as pd
from kivy.clock import Clock
from kivy.lang import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.scrollview import ScrollView
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.button import MDFlatButton
from kivymd.uix.dialog import MDDialog
from kivymd.uix.list import MDList, OneLineListItem
from kivymd.uix.list import ThreeLineListItem
from kivymd.uix.textfield import MDTextField
from plyer import storagepath
from plyer import filechooser

KV = '''
<Manager>:
    Home:
        name: 'home'

    Entries:
        name: 'new'

    #View:
     #   name: 'view'

<Home>:
    NavigationLayout:
        ScreenManager:
            Screen:
                BoxLayout:
                    orientation: 'vertical'
                    MDToolbar:
                        title: 'Home'
                        left_action_items: [["menu", lambda x: nav_draw.set_state()]]
                        right_action_items: [["comment-question", lambda x: None]]
                        elevation: 8
                    MDLabel:

                        text: 'Welcome'
                        #font_style: 'H1'
                        halign: 'center'
                        font_size: 100
                        bold: True

                    MDLabel:
                        id: text
                        font_style: 'H2'
                        halign: 'center'
                        #font_size: 100
                        #bold: True


        MDNavigationDrawer:
            id: nav_draw
            BoxLayout:
                orientation: 'vertical'

                ScrollView:
                    MDList:
                        OneLineIconListItem:
                            text: "Add Entries"
                            on_press:
                                nav_draw.set_state('close')
                                root.manager.current = 'new'

                            IconLeftWidget:
                                icon: 'file-document-box-plus'

                        OneLineIconListItem:
                            text: "View Entries"
                            on_press:
                                nav_draw.set_state('close')
                                root.manager.current = 'view'

                            IconLeftWidget:
                                icon: 'book-open'


<Entries>:
    NavigationLayout:
        ScreenManager:
            Screen:
                BoxLayout:
                    orientation: 'vertical'
                    MDToolbar:
                        title: 'Add Entries'
                        left_action_items: [["menu", lambda x: nav_draw.set_state()]]
                        right_action_items: [["comment-question", lambda x: None]]
                        elevation: 8

                    MDFloatLayout:

                        MDTextField:
                            id: reg_no
                            hint_text: 'Enter the registration number here'
                            pos_hint: {'center_x': 0.5, 'center_y': 0.5}
                            size_hint_x: .75
                            on_text_validate: root.check()

                        MDTextField:
                            id: remarks
                            hint_text: "Select Remark"
                            pos_hint: {'center_x': 0.5, 'center_y': 0.3}
                            size_hint_x: .75
                            on_focus:
                                if self.focus: root.issues()
                                root.check()


                        MDFloatingActionButton:
                            icon: 'file-document-box-plus'
                            text: "Done"
                            pos_hint: {'center_x': 0.5, 'center_y': .1}
                            on_release: root.save()
                    Widget:

        MDNavigationDrawer:
            id: nav_draw
            BoxLayout:
                orientation: 'vertical'

                ScrollView:
                    MDList:
                        OneLineIconListItem:
                            text: "Home"
                            on_press:
                                nav_draw.set_state('close')
                                root.manager.current = 'home'

                            IconLeftWidget:
                                icon: 'home'

                        OneLineIconListItem:
                            text: "View Entries"
                            on_press:
                                nav_draw.set_state('close')
                                root.manager.current = 'view'

                        OneLineIconListItem:
                            text: "New"
                            on_press:
                                nav_draw.set_state('close')
                                root.master_loc()

                            IconLeftWidget:
                                icon: 'book-open'

<View>:
    NavigationLayout:
        ScreenManager:
            Screen:
                BoxLayout:
                    orientation: 'vertical'
                    MDToolbar:
                        title: 'Saved Issues'
                        left_action_items: [["menu", lambda x: nav_draw.set_state()]]
                        right_action_items: [["comment-question", lambda x: None]]
                        elevation: 8

                    ScrollView:
                        MDList:
                            id: lister

        MDNavigationDrawer:
            id: nav_draw
            BoxLayout:
                orientation: 'vertical'

                ScrollView:
                    MDList:
                        OneLineIconListItem:
                            text: "Home"
                            on_press:
                                nav_draw.set_state('close')
                                root.manager.current = 'home'

                            IconLeftWidget:
                                icon: 'home'

                        OneLineIconListItem:
                            text: "Add Entries"
                            on_press:
                                nav_draw.set_state('close')
                                root.manager.current = 'new'

                            IconLeftWidget:
                                icon: 'file-document-box-plus'

'''

slave = ""
master = ""
name = ""


class Home(Screen):
    def on_enter(self, *args):
        if master:
            pass
        else:
            Clock.schedule_once(self.master_loc, .5)
        self.clear = False

    def load_entry_no(self, *args):
        if self.clear is True:
            lis = download(slave)
            self.ids.text.text = f"You have {len(lis)} entries"

    def loc(self, *args):

        self.dialog = MDDialog(title="Edit", text="HI Do you want to load or create a new file to save from?",
                               buttons=[MDFlatButton(text='New', on_press=self.new),
                                        MDFlatButton(text='Load', on_press=self.load),
                                        MDFlatButton(text='Cancel', on_press=self.close),
                                        ],
                               )
        self.dialog.set_normal_height()
        self.dialog.open()

    def master_loc(self, *args):
        self.dialog1 = MDDialog(title="Edit", text="HI please select the file you are using",
                                buttons=[MDFlatButton(text='Okay', on_press=self.load_master),
                                         ],
                                )
        self.dialog1.set_normal_height()
        self.dialog1.open()

    def load_master(self, *args):
        z = storagepath.get_documents_dir()
        m = filechooser.open_file()
        get_master(m)
        self.dialog1.dismiss()

        self.loc()

    def load(self, *args):
        self.close()
        self.location = filechooser.open_file()
        get_slave(self.location)

    def new(self, *args):
        self.close()
        layout = MDBoxLayout(orientation='vertical', spacing='12dp', size_hint_y=None, height="20dp")
        self.nam = MDTextField()
        layout.add_widget(self.nam)

        self.dialog = MDDialog(title="Edit", type='custom', content_cls=layout,
                               buttons=[MDFlatButton(text='Okay', on_press=self.ok),
                                        ],
                               )
        # self.dialog.set_normal_height()
        self.dialog.open()

    def ok(self, *args):
        if self.nam.text:
            s = self.nam.text.title()
            get_name(s)

        self.close()

    def close(self, *args):
        self.dialog.dismiss()
        self.dialog1.dismiss()


def download(loc):
    cf = pd.read_excel(loc, engine='openpyxl')
    df = cf.to_dict('records')
    return df


def upload_1(loc, lis):
    df = pd.DataFrame(lis)
    df.to_excel(loc, index=False)


def upload_2(loc, lis):
    df = pd.DataFrame.from_dict(lis)
    df.to_excel(loc, index=False)


def save(lis, name):
    z = storagepath.get_documents_dir()

    if os.path.exists(name[0]):

        upload_2(name[0], lis)
    else:
        os.mkdir(f"{z}/Helper Export/")
        upload_1(f"{z}/Helper Export/{name}.xlsx", lis)


def get_slave(value):
    global slave
    slave = value[0]


def get_master(value):
    global master
    master = value


def get_name(value):
    global name
    name = value


class Entries(Screen):
    def on_enter(self, *args):

        if master:
            self.df = download(master[0])
        else:
            self.master_loc()
        self.target = ""
        self.others = False
        self.target = ""
        self.others = False
        self.found = False

    def loc(self, *args):

        self.dialog = MDDialog(title="Edit", text="HI Do you want to load or create a new file to save from?",
                               buttons=[MDFlatButton(text='New', on_press=self.new),
                                        MDFlatButton(text='Load', on_press=self.load),
                                        MDFlatButton(text='Cancel', on_press=self.close),
                                        ],
                               )
        self.dialog.set_normal_height()
        self.dialog.open()

    def master_loc(self, *args):
        self.dialog1 = MDDialog(title="Edit", text="HI please select the file you are using",
                                buttons=[MDFlatButton(text='Okay', on_press=self.load_master),
                                         ],
                                )
        self.dialog1.set_normal_height()
        self.dialog1.open()

    def load_master(self, *args):
        z = storagepath.get_documents_dir()
        m = filechooser.open_file()
        get_master(m)
        self.df = download(master[0])
        self.dialog1.dismiss()

        self.loc()

    def load(self, *args):
        self.close()
        self.location = filechooser.open_file()
        get_slave(self.location)

    def new(self, *args):
        self.close()
        layout = MDBoxLayout(orientation='vertical', spacing='12dp', size_hint_y=None, height="20dp")
        self.nam = MDTextField()
        layout.add_widget(self.nam)

        self.dialog = MDDialog(title="Edit", type='custom', content_cls=layout,
                               buttons=[MDFlatButton(text='Okay', on_press=self.ok),
                                        ],
                               )
        # self.dialog.set_normal_height()
        self.dialog.open()

    def ok(self, *args):
        if self.nam.text:
            s = self.nam.text.title()
            get_name(s)
        self.close()

    def check(self):
        for x in range(len(self.df)):
            if self.df[x]['REGNO'] == self.ids.reg_no.text:
                self.target = x
                self.found = True
        if self.found is False:
            self.warning = MDDialog(title="Warning", text="Registration Number Not Found",
                                    buttons=[MDFlatButton(text="CLose")])
            self.warning.open()

    def issues(self):
        if self.others is False:
            frame = MDBoxLayout(orientation='vertical', size_hint_y=None, height=150)
            scroll = ScrollView()
            list_view = MDList()

            scroll.add_widget(list_view)

            remarks = ['Awaiting Result', 'Token not found', 'Passport Uploaded', 'Wrong result uploaded',
                       'No results uploaded', 'Card limit exceeded', 'Invalid card', 'Result not uploaded',
                       'Incomplete Result', 'Result not visible', 'Invalid pin', 'Invalid serial',
                       'Result checker has been used', 'Pin for Neco not given',
                       'Wrong result uploaded', 'Incomplete result',
                       'Token linked to another candidate', 'Others']

            for x in remarks:
                list_view.add_widget(OneLineListItem(text=x, on_press=self.get_selection))

            self.chooser = MDDialog(title='Select Remark', size_hint=(.5, .4),
                                    type='custom', content_cls=frame)
            frame.add_widget(scroll)
            # self.chooser.set_normal_height()
            self.chooser.open()

    def get_selection(self, instance):
        if instance.text != 'Others':
            self.ids.remarks.text = instance.text

        else:
            self.others = True
            Clock.schedule_once(self.prompt, 0.1)
        self.chooser.dismiss()

    def prompt(self, *args):
        self.ids.remarks.focus = True
        self.ids.remarks.hint_text = "Type Issue here"

    def save(self):
        if self.ids.reg_no.text and self.ids.remarks.text:
            self.df[int(self.target)]['REMARK'] = self.ids.remarks.text

            if slave:
                if self.found:
                    z = download(slave)
                    self.go = True
                    for v in z:
                        if self.ids.reg_no.text in v['REGNO']:
                            self.go = False
                            break
                    if self.go is True:
                        z.append(self.df[int(self.target)])
                        upload_2(slave, z)
            elif name:

                f = storagepath.get_documents_dir()
                n = []

                n.append(self.df[int(self.target)])
                if os.path.exists(f"{f}/Helper Export/"):

                    upload_1(f"{f}/Helper Export/{name}.xlsx", n)
                else:
                    os.mkdir(f"{f}/Helper Export/")

                    upload_1(f"{f}/Helper Export/{name}.xlsx", n)
                get_slave([f"{f}/Helper Export/{name}.xlsx"])

            self.others = False
            self.ids.remarks.hint_text = "Select Remark"
            self.ids.reg_no.text = ""
            self.ids.remarks.text = ""

        else:
            self.warning = MDDialog(title="Warning", text="Please Fill All Values",
                                    buttons=[MDFlatButton(text="CLose")])
            self.warning.open()

    def close(self, *args):
        self.dialog.dismiss()
        self.dialog1.dismiss()


class View(Screen):
    class View(Screen):
        def on_enter(self, *args):
            self.reload()
            if slave:

                self.lis = download(slave)
            else:
                self.lis = []

        def load(self):

            for x in range(len(self.lis)):
                items = ThreeLineListItem(text=f'No {x + 1}', secondary_text=f"REGNO: {self.lis[x]['REGNO']}",
                                          tertiary_text=f"Issue: {self.lis[x]['REMARK']}",

                                          on_release=self.target)
                self.ids.lister.add_widget(items)

        def target(self, instance):
            """Locates the index no in the list"""
            lis = self.lis
            for x in range(len(lis)):
                if lis[x]['REGNO'] == instance.secondary_text.lstrip('REGNO: '):
                    self.edit()
                    self.reg.text = lis[x]['REGNO']
                    self.issue.text = lis[x]['REMARK']
                    self.target_no = x

        def edit(self, *args):
            layout = MDBoxLayout(orientation='vertical', spacing='12dp', size_hint_y=None, height="120dp")
            self.reg = MDTextField()
            self.issue = MDTextField(multiline=True)
            layout.add_widget(self.reg)
            layout.add_widget(self.issue)
            self.dialog = MDDialog(title="Edit", type='custom', content_cls=layout,
                                   buttons=[MDFlatButton(text='Save', on_press=self.save),
                                            MDFlatButton(text='Delete', on_press=self.delete_prompter),
                                            MDFlatButton(text='Cancel', on_press=self.close),
                                            ],
                                   )
            self.dialog.set_normal_height()
            self.dialog.open()

        def save(self, *args):
            pass

        def saveR(self, *args):

            # a list for the values in the dialog box
            edit = []
            # downloads the current excel file to lis

            # gets the values in the dialog box
            for obj in self.dialog.content_cls.children:
                if isinstance(obj, MDTextField):
                    edit.append(obj.text)

            # changes the value of the current target to the new ones edited
            self.lis[int(self.target_no)]['REGNO'] = edit[1].upper()
            self.lis[int(self.target_no)]['REMARK'] = edit[0].title()
            upload_2(slave, self.lis)
            self.dialog.dismiss()
            self.reload()

        def close(self, *args):
            self.dialog.dismiss()

        def reload(self):
            """Refresh this page"""
            self.ids.lister.clear_widgets()
            self.lis = []
            self.lis = download(slave)
            self.load()

        def delete_prompter(self, *args):
            """Prevents the user from accidentally deleting one of the list items"""

            text = f"Are you sure you want to delete the entry {self.lis[int(self.target_no)]['REGNO']} with the issue of " \
                   f"'{self.lis[int(self.target_no)]['REMARK']}'?"
            self.dialog.dismiss()
            self.warning(text)

        def delete(self, *args):
            # lis = download()
            self.lis.pop(int(self.target_no))
            upload_2(slave, self.lis)
            self.dialog.dismiss()
            self.reload()

        def warning(self, text):
            self.dialog = MDDialog(title="Edit", text=text,
                                   buttons=[MDFlatButton(text='Yes', on_press=self.delete),
                                            MDFlatButton(text='Cancel', on_press=self.close),
                                            ],
                                   )
            self.dialog.set_normal_height()
            self.dialog.open()


class Manager(ScreenManager):
    pass


class Helper(MDApp):
    def __init__(self, **kwargs):
        super(Helper, self).__init__(**kwargs)
        Builder.load_string(KV)

    def build(self):
        self.theme_cls.theme_style = "Dark"
        # self.theme_cls.font_style = 'Gray'
        return Manager()


Helper().run()