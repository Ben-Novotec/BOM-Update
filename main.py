from kivymd.app import MDApp
from plyer import filechooser
import bom_update


class MainApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.bom_old_path = ''
        self.bom_new_path = ''

    def build(self):
        self.theme_cls.theme_style = 'Dark'
        self.theme_cls.primary_palette = 'Blue'
        pass

    def set_bom_old_path(self):
        self.bom_old_path = filechooser.open_file(title='choose single component BOM list',
                                                  filters=[('Excel file', '*.xlsx')])[0]
        self.root.ids.lab_old_path.text = self.bom_old_path

    def set_bom_new_path(self):
        self.bom_new_path = filechooser.open_file(title='choose hierarchical BOM list',
                                                  filters=[('Excel file', '*.xlsx')])[0]
        self.root.ids.lab_new_path.text = self.bom_new_path

    def update_bom(self):
        bom_update.main(self.bom_old_path, self.bom_new_path)
        self.root.ids.run_button.text = 'BOM list updated!'


if __name__ == '__main__':
    MainApp().run()
