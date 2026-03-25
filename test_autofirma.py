import subprocess

input_pdf = "Plantilla Compensación Gastos ABR 2026.pdf"
output_pdf = "test_signed.pdf"

cmd = ["AutoFirma", "commandline", "-i", input_pdf, "-o", output_pdf, "-format", "pdf", "-store", "auto"]
print("Running:", " ".join(cmd))
try:
    result = subprocess.run(cmd, capture_output=True, text=True) # Note: NOT using CREATE_NO_WINDOW
    print("Return code:", result.returncode)
    print("STDOUT:", result.stdout)
    print("STDERR:", result.stderr)
except Exception as e:
    print("Exception running AutoFirma:", e)
