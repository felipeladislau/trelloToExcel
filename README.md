# Trello to Excel
Processa o JSON do trello e exporta os cards para Excel

O primeiro passo é substituir os dados de data.json pelos dados obtidos do Trello, para isto acesse o quadro desejado, clique nas opções, em "Imprimir e exportar" e em JSON.

Copie e cole o conteúdo no data.json.

Depois basta rodar o script que o mesmo vai criar um arquivo de excel (XLSX) com os dados do trello.

Nesta versão o arquivo criado contém apenas a data e hora da criação do card, nome, descrição e labels, porém o array $cards contém todos os dados de cada card, portanto é possível adicionar mais dados.

Para isto basta incluir após a linha 31 (No final do array) o dado extra que deseja adicionar e após isto na linha 45 o nome da coluna.

Por fim adicione na linha 58 o dado com base no array.
