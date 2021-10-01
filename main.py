from kivymd.app import MDApp
from plyer import filechooser
import bom_update


class MainApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.bom_single_path = ''
        self.bom_hier_path = ''

    def build(self):
        self.theme_cls.theme_style = 'Dark'
        self.theme_cls.primary_palette = 'Blue'
        pass

    def set_bom_single_path(self):
        self.bom_single_path = filechooser.open_file(title='choose single component BOM list',
                                                     filters=[('Excel file', '*.xlsx')])[0]
        self.root.ids.lab_single_path.text = self.bom_single_path

    def set_bom_hier_path(self):
        self.bom_hier_path = filechooser.open_file(title='choose hierarchical BOM list',
                                                   filters=[('Excel file', '*.xlsx')])[0]
        self.root.ids.lab_hier_path.text = self.bom_hier_path

    def update_bom(self):
        bom_update.main(self.bom_single_path, self.bom_hier_path)
        self.root.ids.run_button.text = 'BOM list updated!'


if __name__ == '__main__':
    MainApp().run()
