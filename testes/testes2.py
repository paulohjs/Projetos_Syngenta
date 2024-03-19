import os

# Defina o caminho para a pasta
folder_path = r'C:\Users\s1002322\Syngenta\Brazil HROps - Comp Cycle - Zip_SIP'

# Percorre a árvore de diretórios
for root, dirs, files in os.walk(folder_path):
    for filename in files:
        # Constrói o caminho completo para o arquivo
        file_path = os.path.join(root, filename)
        print(file_path) # Isso imprimirá o caminho completo do arquivo



# # Dados fornecidos
# dados = """
# manager_name	email	nome_arquivo
# Adriano Mittmann	adriano.mittmann@syngenta.com	Adriano Mittmann.zip
# Aimar Pedrini	aimar.pedrini@syngenta.com	Aimar Pedrini.zip
# Alex Federich	alex.federich@syngenta.com	Alex Federich.zip
# Alexandre Lellis	alexandre.lellis@syngenta.com	Alexandre Lellis.zip
# Alexandre Marineli	alexandre.marineli@syngenta.com	Alexandre Marineli.zip
# Alexandre Monteiro	alexandre.monteiro@syngenta.com	Alexandre Monteiro.zip
# Andre Nunes	andre.nunes@syngenta.com	Andre Nunes.zip
# Ariovaldo Filho	ariovaldo.quintana@syngenta.com	Ariovaldo Filho.zip
# Ather Barros	ather.barros@syngenta.com	Ather Barros.zip
# Benigno Almeida	Benigno.Almeida@syngenta.com	Benigno Almeida.zip
# Bruno Franco	bruno.franco@syngenta.com	Bruno Franco.zip
# Bruno Miranda	bruno.miranda@syngenta.com	Bruno Miranda.zip
# Bruno Takay	bruno.takay@syngenta.com	Bruno Takay.zip
# Carlos Tonet	carlos.tonet@syngenta.com	Carlos Tonet.zip
# Chrismam Mrozinski	Chrismam.Mrozinski@syngenta.com	Chrismam Mrozinski.zip
# Claudionor Filho	Claudionor.Filho@syngenta.com	Claudionor Filho.zip
# Cleiton Antonio da Silva Barbosa	Cleiton.Barbosa@syngenta.com	Cleiton Antonio da Silva Barbosa.zip
# Cristiane Martin	Cristiane.Martin@syngenta.com	Cristiane Martin.zip
# Daniel Stuchi	daniel.stuchi@syngenta.com	Daniel Stuchi.zip
# Danilo Cestari	danilo.cestari@syngenta.com	Danilo Cestari.zip
# Danilo Goncalves	danilo.goncalves@syngenta.com	Danilo Goncalves.zip
# Douglas Silva	douglas.silva@syngenta.com	Douglas Silva.zip
# Dylnei Neto	Dylnei.Neto@syngenta.com	Dylnei Neto.zip
# Edgar Garcia	Edgar.Garcia@syngenta.com	Edgar Garcia.zip
# Ediney Furtado	ediney.furtado@syngenta.com	Ediney Furtado.zip
# Edison Peron	edison.peron@syngenta.com	Edison Peron.zip
# Emerson Barbizan	Emerson.Barbizan@syngenta.com	Emerson Barbizan.zip
# Fabiano Argenta	Fabiano.Argenta@syngenta.com	Fabiano Argenta.zip
# Felipe Fett	felipe.fett@syngenta.com	Felipe Fett.zip
# Felipe Manerba	felipe.manerba@syngenta.com	Felipe Manerba.zip
# Fernando Bertolo	fernando.bertolo@syngenta.com	Fernando Bertolo.zip
# Francis Palotta	francis.palotta@syngenta.com	Francis Palotta.zip
# Gerhard Vitor Ecker	Gerhard.Ecker@syngenta.com	Gerhard Vitor Ecker.7z
# Gilberto Landgraf	gilberto.landgraf@syngenta.com	Gilberto Landgraf.7z
# Gilberto Piza	gilberto.piza@syngenta.com	Gilberto Piza.7z
# Guilherme Adami	Guilherme.Adami@syngenta.com	Guilherme Adami.7z
# Gustavo Antunes	Gustavo.Antunes@syngenta.com	Gustavo Antunes.7z
# Gustavo Freitas	gustavo.freitas@syngenta.com	Gustavo Freitas.7z
# Gustavo Horta	gustavo.horta@syngenta.com	Gustavo Horta.7z
# Henrique Mourao	henrique.mourao@syngenta.com	Henrique Mourao.7z
# Humberto Rosada	humberto.rosada@syngenta.com	Humberto Rosada.7z
# Isnard Campos	Isnard.Campos@syngenta.com	Isnard Campos.7z
# Jairo Oliveira	jairo.oliveira@syngenta.com	Jairo Oliveira.7z
# Joao Lima	Joao.Lima@syngenta.com	Joao Lima.7z
# Joao Lincoln	Joao.Lincoln@syngenta.com	Joao Lincoln.7z
# Jocimar Mauri	jocimar.mauri@syngenta.com	Jocimar Mauri.7z
# Jonas Gonzatti	jonas.gonzatti@syngenta.com	Jonas Gonzatti.7z
# Jonatha Bolzan	jonatha.bolzan@syngenta.com	Jonatha Bolzan.7z
# Jose Barao	jose.eduardo@syngenta.com	Jose Barao.7z
# Juliano Castaldo	juliano.castaldo@syngenta.com	Juliano Castaldo.7z
# Kanoyo Franco	Kanoyo.Franco@syngenta.com	Kanoyo Franco.7z
# Leonardo Asseff Velasquez	leonardo.velasquez@syngenta.com	Leonardo Asseff Velasquez.7z
# Leopoldo Diniz	leopoldo.diniz@syngenta.com	Leopoldo Diniz.7z
# Lucas Lamounier	Lucas.Lamounier@syngenta.com	Lucas Lamounier.7z
# Lucas Tonello	lucas.tonello@syngenta.com	Lucas Tonello.7z
# Lucio Zabot	lucio.zabot@syngenta.com	Lucio Zabot.7z
# Luis Gustavo Silva	luis_gustavo.neves@syngenta.com	Luis Gustavo Silva.7z
# Luiz Oliveira	luiz.oliveira-1@syngenta.com	Luiz Oliveira.7z
# Marcelo Goncalves	marcelo.goncalves@syngenta.com	Marcelo Goncalves.7z
# Marcelo Pansera	Marcelo.Pansera@syngenta.com	Marcelo Pansera.7z
# Marcos Basso	marcos.basso@syngenta.com	Marcos Basso.7z
# Marcos Kohlrausch	marcos.kohlrausch@syngenta.com	Marcos Kohlrausch.7z
# Mariana Moises Goulart	Mariana_Moises.Goulart@syngenta.com	Mariana Moises Goulart.7z
# Mario Isquierdo	mario.isquierdo@syngenta.com	Mario Isquierdo.7z
# Matheus Alvim	matheus.alvim@syngenta.com	Matheus Alvim.7z
# Matheus Pedroni	matheus.pedroni@syngenta.com	Matheus Pedroni.7z
# Paulo Arruda	paulo.arruda@syngenta.com	Paulo Arruda.7z
# Pedro Scarton	pedro.scarton@syngenta.com	Pedro Scarton.7z
# Querson Fornari	querson_roberto.fornari@syngenta.com	Querson Fornari.7z
# Rafael Andrade	rafael.andrade@syngenta.com	Rafael Andrade.7z
# Rafael Chioquetta	rafael.chioquetta@syngenta.com	Rafael Chioquetta.7z
# Rafael Pereira Oliveira	rafael_pereira.oliveira@syngenta.com	Rafael Pereira Oliveira.7z
# Rafael Ribeiro	Rafael_Ferreira.Ribeiro@syngenta.com	Rafael Ribeiro.7z
# Ramirez Toledo	Ramirez.Toledo@syngenta.com	Ramirez Toledo.7z
# Renan Krug	renan.krug@syngenta.com	Renan Krug.7z
# Renato Cardoso	Renato.Cardoso@syngenta.com	Renato Cardoso.7z
# Rodolfo Bomfim	rodolfo.bomfim@syngenta.com	Rodolfo Bomfim.7z
# Rodrigo Alves	rodrigo.alves@syngenta.com	Rodrigo Alves.7z
# Rodrigo Dourado	Rodrigo.Dourado@syngenta.com	Rodrigo Dourado.7z
# Rodrigo Pereira	rodrigo.pereira@syngenta.com	Rodrigo Pereira.7z
# Rogerio Ramos	rogerio.ramos@syngenta.com	Rogerio Ramos.7z
# Rudimar Tonin	Rudimar.Tonin@syngenta.com	Rudimar Tonin.7z
# Taffareu Agostineti	Taffareu.Agostineti@syngenta.com	Taffareu Agostineti.7z
# Tales Lima	tales.lima@syngenta.com	Tales Lima.7z
# Talles Lopes Tolentino	talles_lopes.tolentino@syngenta.com	Talles Lopes Tolentino.7z
# Thiago Audi Gimenes	thiago.gimenes@syngenta.com	Thiago Audi Gimenes.7z
# Tulio Santos	tulio.santos@syngenta.com	Tulio Santos.7z
# Tulio Silva	tulio.teodoro@syngenta.com	Tulio Silva.7z
# Valter Filho	valter.soares@syngenta.com	Valter Filho.7z
# Victor Silveira	victor.silveira@syngenta.com	Victor Silveira.7z
# Vinicius Moraes	vinicius.moraes@syngenta.com	Vinicius Moraes.7z
# Vinicius Pianaro	Vinicius.Pianaro@syngenta.com	Vinicius Pianaro.7z
# Vinicius Uzae	Vinicius.Uzae@syngenta.com	Vinicius Uzae.7z
# Welder Fuzita	welder.fuzita@syngenta.com	Welder Fuzita.7z
# Wilian Manoel	wilian.manoel@syngenta.com	Wilian Manoel.7z
# Willie Cintra	willie.cintra@syngenta.com	Willie Cintra.7z

# """

# # Dividindo os dados em linhas
# linhas = dados.strip().split('\n')

# # Ignorando o cabeçalho
# linhas = linhas[1:]

# # Criando o dicionário
# dicionario_managers = {}

# for linha in linhas:
#     partes = linha.split('\t')  # Supondo que o separador seja o caractere de tabulação
#     manager_name, email, nome_arquivo = partes
#     dicionario_managers[manager_name] = {
#         'email': email,
#         'nome_arquivo': nome_arquivo
#     }

# # Imprimindo o dicionário criado

# print(dicionario_managers)
