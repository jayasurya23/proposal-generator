import PyInstaller.__main__
import os

# The name of your main python script
script_name = "proposal_generator.py"

# The name for your final executable
exe_name = "ProposalGenerator"

# Define the data files to be included (logo and fonts)
# The format is 'source_path;destination_in_bundle'
# '.' means the root directory of the bundled app.
data_to_add = [
    'logo.png',
    'Jost-Regular.ttf',
    'Jost-Bold.ttf'
]

# Construct the --add-data arguments for PyInstaller
# The separator is different for Windows (;) vs. Mac/Linux (:)
separator = ';' if os.name == 'nt' else ':'
add_data_args = []
for item in data_to_add:
    add_data_args.append('--add-data')
    add_data_args.append(f'{item}{separator}.')

# --- PyInstaller Command Arguments ---
pyinstaller_args = [
    script_name,
    '--noconfirm',      # Don't ask for confirmation to overwrite
    '--onefile',        # Create a single .exe file
    '--windowed',       # Don't show a console window when the app runs
    f'--name={exe_name}', # Set the name of the executable
]

# Add the data file arguments to the main command
pyinstaller_args.extend(add_data_args)

# --- Run PyInstaller ---
if __name__ == '__main__':
    print("Running PyInstaller with the following arguments:")
    print(' '.join(pyinstaller_args))
    PyInstaller.__main__.run(pyinstaller_args)
    print("\nBuild complete.")
