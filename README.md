# desafioRPA
Automatização de leitura de xml para transcrição em excel com python - RPA

## Code_excelv2

  importação de bibliotecas úteis:
    pandas para criação de dataframe e utilização do excelwriter
    os para navegação nos diretórios
    
   utilização de módulos:
    ET para leitura dos xml e localização das tags
    
   O código perscruta todo o diretório em busca dos arquivos xml, armazena as str em uma lista, transforma-a em um objeto e então o data frame criado é passado para o excel.
   
   tentativa de overlay no modo append não funciona, mesmo após remoção da Data Validation no excel, com programa não iniciando a inserção de dados na linha correta
