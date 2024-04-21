from cx_Freeze import setup, Executable

executables = [Executable(script='Automação.py')]

options = {
    'build_exe': {
'include_files': ['C:\\Users\\claudemir.santos\\Desktop\\V2\\AtualizacaoPlanilhas.xlsx']
    }
}

setup(
    name='Atualização Planilhas Zendesk',
    version='3.0',
    description='Descrição do seu programa',
    options=options,
    executables=executables
)

#python setup.py clean
#python setup.py build