<?php 

        //Recuperando dados
        $ob_relacao_dados = new  OB_DADOS($system);
        $dados = $ob_relacao_dados->GetDados($pessoaID);

        //Gerando excel
        include_once('GerandoExcel.php');

        $rand = rand(0, 100000);
        $arquivo = 'relatorio-dados-' . $pessoaID . $rand . '.xls';
        $excelRelatorio = new EXCEL_RELATORIOS($system);
        $excel = $excelRelatorio->getExcel($faturas, $arquivo);
