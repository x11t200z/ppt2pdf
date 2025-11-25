import os
import comtypes.client
# dùng để tạo đối tượng COM, cho phép python thao tác với PowerPoint như khi người dùng tự thao tác

ppt_dir = r"C:\path\to\ppt" # đường dẫn tới thư mục chứa các file PowerPoint
output_dir = r"C:\path\to\pdf\save" # đường dẫn tới thư mục lưu file PDF

powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
powerpoint.Visible = 1 
# = 1 để hiển thị cửa sổ PowerPoint trong quá trình chuyển đổi, nếu không muốn hiển thị (kiểu như chạy ngầm) thì để = 0

for file in os.listdir(ppt_dir):
    if file.endswith(".ppt") or file.endswith(".pptx"):
        full_path = os.path.join(ppt_dir, file)
        pdf_path = os.path.join(output_dir, file.rsplit(".", 1)[0] + ".pdf")
        presentation = powerpoint.Presentations.Open(full_path)
        presentation.SaveAs(pdf_path, 32)  # 32 = PDF
        presentation.Close()
powerpoint.Quit()
