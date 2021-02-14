from win32com.client import GetActiveObject
import win32com.client as win32

class AIFileHandler:
    def __init__(self, files_path, base_path):
        self.app = GetActiveObject("Illustrator.Application")
        self.files_path = files_path
        self.base_path = base_path

    def process_file(self, file_name):
        print(f'Processing file {file_name}')
        doc = self.app.Open(f'{self.files_path}/{file_name}')
        layers = doc.Layers
        print(f'Total layers: {len(layers)}')

        png_export_options = win32.Dispatch("Illustrator.ExportOptionsPNG24")
        png_export_options.horizontalScale = 200
        png_export_options.verticalScale = 200 

        image_capture_options = win32.Dispatch("Illustrator.imageCaptureOptions")
        image_capture_options.resolution = 250
        image_capture_options.transparency = True
        image_capture_options.antiAliasing = True

        svg_export_options = win32.Dispatch("Illustrator.ExportOptionsSVG")
        svg_export_options.fontType = 2
        svg_export_options.embedRasterImages = True
        # svg_export_options.cssProperties = 3
        svg_export_options.fontSubsetting = 1 # aiNoFonts

        for layer in layers:
            layer.visible = False

        for layer in layers:
            print(f'Processing layer {layer.name}...')
            layer.visible = True
            if len(layer.TextFrames) != 1 or len(layer.GroupItems) != 1:
                print(f'❗ Fail to export layer {layer.name}')
                layer.visible = False
                continue
            text_frame = layer.TextFrames[0]
            exported_name = text_frame.contents
            text_frame.selected = True
            self.app.ExecuteMenuCommand("hide")
            text_frame.selected = False
            group = layer.GroupItems[0]
            bounds = group.controlBounds

            group.selected = True
            # doc.Export(ExportFile=f'{exported_name}.png',ExportFormat=5,Options=png_export_options)
            doc.ImageCapture(f'{self.base_path}/{exported_name}.png', group.controlBounds, image_capture_options)
            doc.FitArtboardToSelectedArt(0)
            doc.Export(ExportFile=f'{self.base_path}/{exported_name}.svg',ExportFormat=3, Options=svg_export_options)
            group.selected = False
            layer.visible = False
            print(f'✓ Succeed to export layer {layer.name}')
            # break
        
        print(f'Exported file {file_name}\r\n')
