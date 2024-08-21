<!-- 
Tutorial de como fazer uma planilha Excel apartir de um banco de dados / Inputs de campos(o codigo daqui é apartir do banco,
mas é facil colocar como se fosse inputs)

Pode se usar duas bibliotecas para isso 
PHPExcel(mais antiga)
e
PHPSheadSheet(mais nova)

aqui iremos utilizar a PHPExcel, lembrando que aqui é apenas uma parte de um codigo muito maior,então algumas coisas interessantes a se acrescentar

-aqui é direto, recomendase ter algum botão em alguma pagina para chamar esse codigo,pode-se usar jQuery(AJAX) 
-Para apagar algo dentro de uma pagina apartir de um programa PHP utilizasse         

unlink($nomeArquivo);

*E existem 3 maneiras para se baixar algo(colocar o simbolozinho em cima da tela)
-apartir de um botão dentro de um php
<a href='" + data +"' download='"nomeArquivo"'><span class = 'btn btn-md'>Download</span></a>

javaScript,depois de um retorno 

window.open(nomeArquivo);
window.location.href=nomeArquivo;
-->

<?php
// Iniciando as variáveis de base, como a conexão com o banco e a exibição de erros
error_reporting(E_ALL & ~E_DEPRECATED);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
require("conecta.php");
require_once dirname(__FILE__) . 'coloca o caminho até PHPExcel';

// Objeto da conexão ao banco
$bd = conecta();

// Você está criando um objeto para manipular a planilha 
$obj = new PHPExcel();

// Essa linha fala para o PHP começar a manipular a partir da planilha inicial 0
$obj->setActiveSheetIndex(0);

// Aqui pega a planilha sendo escrita a início e a deixa como uma variável
$sheet = $obj->getActiveSheet();

// Iniciando variáveis que iremos utilizar posteriormente 
$cidade = "Colombo";
$indicePlanilhaAtual = 0;
$nomeBibliotecaAtual = '';
$i = 1;

// Aqui é basicamente o CSS da planilha Excel, você cria variáveis com as cores/formatações desejadas
$bordas = array(
    'borders' => array(
        'allborders' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN
        )
    )
);

$centralizar = array(
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
    )
);

$titulo = array(
    'borders' => array(
        'allborders' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN
        )
    ),
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => 'dbead5'), // Cor de fundo
    ),
    'font' => array(
        'bold' => true,
        'color' => array('rgb' => '000000') // Cor do texto
    ),
);

// Esse programa pega dados de um banco e insere os dados em uma planilha, então aqui é a query para pegar os dados que queremos 
$query  = "SELECT lb.cod_biblioteca, lb.cod_livro, l.titulo, l.ano_publicacao, a.nome_autor, g.genero_literario, b.nome_biblioteca, b.endereco ";
$query .= "FROM LIVROS_BIBLIOTECA lb ";
$query .= "JOIN LIVROS l ON lb.cod_livro = l.cod_livro ";
$query .= "LEFT JOIN AUTORES_LIVROS al ON l.cod_livro = al.cod_livro ";
$query .= "LEFT JOIN AUTORES a ON al.cod_autor = a.cod_autor ";
$query .= "LEFT JOIN GENEROS_LITERARIOS g ON l.cod_genero = g.cod_genero ";
$query .= "LEFT JOIN BIBLIOTECAS b ON b.cod_biblioteca = lb.cod_biblioteca ";
$query .= "WHERE b.COD_CIDADE = " . Troca($cidade, 1) . " ";
$query .= "ORDER BY b.nome_biblioteca, l.ano_publicacao DESC, a.nome_autor";

/*
Aqui executa a query e verifica se ela retornou um dado. Se for TRUE, continua o código; se não, o código é finalizado.
*/

$bContinua = false;
if ($bd->SqlExecuteQuery($query) && $bd->SqlNumRows() > 0) {
    $bContinua = true;
}

// Enquanto bContinua for true, vai rodando
while ($bContinua) {

    // mesma coisa de cima
    $sheet = $obj->getActiveSheet();
    
    //Centraliza todos os campos 
    $sheet->getDefaultStyle()->applyFromArray($centralizar);

    //pega dois dados do banco e coloca em uma variavel.
    $nomeBiblioteca = $bd->SqlQueryShow('nome_biblioteca');
    $enderecoTexto = $bd->SqlQueryShow('endereco');

   //como são varias bibliotecas,faz este if
    if ($nomeBiblioteca !== $nomeBibliotecaAtual) {
     
      if ($indicePlanilhaAtual > 0) {
         // Define o título da planilha atual com o nome da biblioteca
         $obj->getActiveSheet()->setTitle($nomeBibliotecaAtual);
      }
      
      if ($indicePlanilhaAtual > 0) {
         // Cria uma nova planilha, mas apenas se não for a primeira
         $obj->createSheet();
      }
      
      $obj->setActiveSheetIndex($indicePlanilhaAtual);
      // Define a planilha atual como a ativa
      $sheet = $obj->getActiveSheet();
      // Obtém a planilha ativa para manipulação
      
      $i = 1;
      // Inicia a contagem de linhas na planilha
      
      $obj->getActiveSheet()->getStyle('A3:E3')->applyFromArray($titulo);
      // Aplica o estilo do título nas células A3 até E3

      $sheet->getStyle('A2')->getAlignment()->setWrapText(true);
      // Ajusta o texto da célula A2 para quebra automática
      
      $sheet->mergeCells('A1:E1');
      // Mescla as células de A1 até E1
      
      $sheet->getStyle('A1:E1')->applyFromArray($titulo);
      // Aplica o estilo do título na área mesclada
      
      $sheet->mergeCells('A2:E2');
      // Mescla as células de A2 até E2
      
      $sheet->setCellValue('A1', 'Biblioteca: ' . $nomeBiblioteca)
            ->setCellValue('A2', $enderecoTexto);
      // Define o valor das células A1 e A2 com o nome da biblioteca e o endereço
      
      $obj->getActiveSheet()->getStyle('A1:E2')->applyFromArray($bordas);
      // Aplica bordas nas células A1 até E2
      
      $sheet->getRowDimension(2)->setRowHeight(45);
      // Define a altura da linha 2 para 45 pixels
      
      foreach (range('A', 'E') as $colunas) {
         $sheet->getColumnDimension($colunas)->setAutoSize(true);
         // Ajusta automaticamente a largura das colunas de A a E
      }
      
      $i += 2;
      // Incrementa a contagem de linhas
      
      $sheet->setCellValue('A' . $i, 'Título do Livro')
            ->setCellValue('B' . $i, 'Ano de Publicação')
            ->setCellValue('C' . $i, 'Autor')
            ->setCellValue('D' . $i, 'Gênero Literário')
            ->setCellValue('E' . $i, 'Localização na Biblioteca');
      // Preenche os cabeçalhos das colunas com os títulos respectivos
      
      $obj->getActiveSheet()->getStyle('A' . $i . ':E' . $i)->applyFromArray($bordas);
      // Aplica bordas nas células da linha atual
      
      $i++;
      // Incrementa a contagem de linhas para a próxima linha de dados
      
      $nomeBibliotecaAtual = $nomeBiblioteca;
      // Atualiza o nome da biblioteca atual
      
      $indicePlanilhaAtual++;
      // Incrementa o índice da planilha para criar a próxima
      
      $tituloLivro = $bd->SqlQueryShow('titulo');
      // Busca o título do livro atual no banco de dados
      
      $sheet->setCellValue('A' . $i, $tituloLivro);
      // Define o valor da célula A na linha atual com o título do livro
      
      $inicio = $i;  
      $final = $i;   
      // Inicializa os índices de início e final para o processo de mesclagem
      
      while ($tituloLivro == $bd->SqlQueryShow('titulo') && $bContinua) {
         // Continua preenchendo as linhas enquanto o título do livro for o mesmo
         $anoPublicacao = $bd->SqlQueryShow('ano_publicacao');
         $nomeAutor = $bd->SqlQueryShow('nome_autor');
         $generoLiterario = $bd->SqlQueryShow('genero_literario');
      
         $sheet->setCellValue('B' . $final, $anoPublicacao)
               ->setCellValue('C' . $final, $nomeAutor)
               ->setCellValue('D' . $final, $generoLiterario);
         // Preenche as colunas com o ano de publicação, autor e gênero literário
      
         $bContinua = $bd->SqlFetchNext();
         // Busca o próximo registro no banco de dados
      
         $final++;
         // Incrementa o índice final para a próxima linha de dados
      }
      
      if ($inicio != $final - 1) {
         // Mescla as células se houver mais de uma linha com o mesmo título
         $obj->getActiveSheet()->mergeCells('A' . $inicio . ':A' . ($final - 1));
      }
      
      $obj->getActiveSheet()->getStyle('A' . $inicio . ':E' . ($final - 1))->applyFromArray($bordas);
      // Aplica bordas nas células mescladas
      
      $i = $final;
      // Atualiza o índice de linhas para a próxima inserção
    } else {
      // Atualiza o valor do título e do índice final
      $inicio = $i;
      $tituloLivro = $bd->SqlQueryShow('titulo');
      $sheet->setCellValue('A' . $i, $tituloLivro);
      $final = $i;

      while ($tituloLivro == $bd->SqlQueryShow('titulo') && $bContinua) {
         // Continua preenchendo as colunas enquanto o título do livro for o mesmo
         $anoPublicacao = $bd->SqlQueryShow('ano_publicacao');
         $nomeAutor = $bd->SqlQueryShow('nome_autor');
         $generoLiterario = $bd->SqlQueryShow('genero_literario');

         $sheet->setCellValue('B' . $final, $anoPublicacao)
               ->setCellValue('C' . $final, $nomeAutor)
               ->setCellValue('D' . $final, $generoLiterario);
         // Preenche as colunas com o ano de publicação, autor e gênero literário

         $bContinua = $bd->SqlFetchNext();
         // Busca o próximo registro no banco de dados

         $final++;
         // Incrementa o índice final para a próxima linha de dados
      }

      if ($inicio != $final - 1) {
         // Mescla as células se houver mais de uma linha com o mesmo título
         $obj->getActiveSheet()->mergeCells('A' . $inicio . ':A' . ($final - 1));
      }

      $obj->getActiveSheet()->getStyle('A' . $inicio . ':E' . ($final - 1))->applyFromArray($bordas);
      // Aplica bordas nas células mescladas

      $i = $final;
      // Atualiza o índice de linhas para a próxima inserção
    }
}

      
$objWriter = PHPExcel_IOFactory::createWriter($obj, 'Excel2007');
// Cria um objeto de escrita para salvar o arquivo no formato Excel 2007

$nomeArquivo = 'planilha/' . gerarNome() . '.xlsx';
// Gera um nome de arquivo único usando uma função chamada gerarNome(); que cria letras aleatorias apartir de:
/*
Em PHP
function gerarNome($tamanho=6){
   $retorno="";
   
   for($i=0;$i<tamanho;$i++){
      if($tamanho%2==0){
      $retorno .= chr(rand(65,90));
      }
      else{
         $retorno .= chr(rand(49,57));
      }
   }
   retorna $retorno;
}

Em JavaScript
function gerarNome(tamanho = 6) {
    let retorno = "";

    for (let i = 0; i < tamanho; i++) {
        if (tamanho % 2 === 0) {
            retorno += String.fromCharCode(Math.floor(Math.random() * (90 - 65 + 1)) + 65);
        } else {
            retorno += String.fromCharCode(Math.floor(Math.random() * (57 - 49 + 1)) + 49);
        }
    }
    return retorno;
}
*/

$objWriter->save($nomeArquivo);
// Salva o arquivo Excel no local especificado

echo $nomeArquivo;
// Exibe o nome do arquivo gerado

$bd->SqlDisconnect();
// Fecha a conexão com o banco de dados
?>
