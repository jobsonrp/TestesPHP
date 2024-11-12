<?php
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\IOFactory;
use Mpdf\Mpdf;

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    // Carrega o template .docx
    $template = new TemplateProcessor('declaracao_template.docx');

    // Substitui os placeholders com os dados do formulário
    $template->setValue('nome', htmlspecialchars($_POST['nome']));
    $template->setValue('cpf', htmlspecialchars($_POST['cpf']));
    $template->setValue('curso', htmlspecialchars($_POST['curso']));

    // Salva o documento preenchido em formato .docx temporário
    $tempDocx = tempnam(sys_get_temp_dir(), 'declaracao') . '.docx';
    $template->saveAs($tempDocx);

    try {
        // Converte o arquivo .docx para HTML simplificado
        $phpWord = IOFactory::load($tempDocx);
        $htmlWriter = IOFactory::createWriter($phpWord, 'HTML');
        
        // Captura o HTML em uma variável
        ob_start();
        $htmlWriter->save('php://output');
        $content = ob_get_clean();

        // Ajustes no HTML para melhorar compatibilidade com mPDF
        // Remove tags que possam causar problemas de layout
        $content = preg_replace('/<style\b[^>]*>(.*?)<\/style>/is', '', $content); // Remove CSS interno
        $content = preg_replace('/<script\b[^>]*>(.*?)<\/script>/is', '', $content); // Remove JavaScript
        $content = preg_replace('/<link\b[^>]*>/is', '', $content); // Remove links externos
        $content = preg_replace('/\{.*?\}/', '', $content); // Remove possíveis caracteres "{" indesejados

        // Criar uma nova instância do mPDF com margens menores
        $mpdf = new Mpdf(['margin_left' => 10, 'margin_right' => 10, 'margin_top' => 10, 'margin_bottom' => 10]);

        // Carrega o conteúdo HTML no mPDF
        $mpdf->WriteHTML($content);
    
        // Define cabeçalho para download do PDF
        $pdfFileName = 'declaracao.pdf';
        $mpdf->Output($pdfFileName, \Mpdf\Output\Destination::DOWNLOAD);
    
        // Limpa o arquivo temporário
        unlink($tempDocx);
    } catch (\Exception $e) {
        // Caso haja algum erro, exiba a mensagem
        echo 'Erro ao gerar PDF: ' . $e->getMessage();
    }
}
?>
