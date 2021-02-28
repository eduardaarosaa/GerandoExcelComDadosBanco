<?php
include_once('modelDados.php');
class EXCEL_RELATORIOS
{
    public function getExcel($faturas, $arquivo)
    {

        $tabela = '<meta charset="UTF-8">';
        $tabela .= '<table>';
        $tabela .= '<tr style="height:80px">';
        $tabela .= '<td style="background-color: #F26522;" colspan="5" h><img src="" width="200px" style="margin-left:15px;" align="center"></td>';
        $tabela .= '<td style="font-weight: bold; font-size:24pt;" colspan="6">Relação de dados</tr>';
        $tabela .= '</tr>';
        $tabela .= '<tr>';
        $tabela .= '<td colspan="20">';
        $tabela .= '</td>';
        $tabela .= '</tr>';
        $tabela .= '<tr>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center; border:none;"><b>Empresa</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Fatura</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Reserva</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Contrato</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Emissão</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Vencimento</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Pagamento</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Vlr. da Fatura</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Vlr. Pago</b></td>';
        $tabela .= '<td colspan="2.5" style="background-color: #F26522; font-size:16pt; text-align:center;"><b>Status</b></td>';
        $tabela .= '</tr>';


        foreach ($dados as $dado){
            $tabela .= '<tr>';
            $tabela .= "<td colspan='2.5' style='text-align: center;'>" . $dado['CNPJ'] ."</td>";
            $tabela .= "<td colspan='2.5' style='text-align: center;'>" . $dado['NroDocumento'] . "</td>";
            $tabela .= "<td colspan='2.5' style='text-align: center;'>" . $dado['ContratoNro'] ."</td>";
            $tabela .= "<td colspan='2.5' style='text-align: center;'>" . $dado['DataEmissao'] ."</td>";
            $tabela .= "<td colspan='2.5' style='text-align: center;'>" . $dado['Vencimento'] ."</td>";
            $tabela .= "<td colspan='2.5'style='text-align: center;'>" . $dado['Pagamento'] ."</td>";
            $tabela .= "<td colspan='2.5' style='text-align: center;'>R$ " . $dado['Total'] ."</td>";
            $tabela .= "<td colspan='2.5' style='text-align: center;'>R$ " . $dado['TotalPago'] ."</td>";
            $tabela .= "<td colspan='2.5' style='text-align: center;'>" . $dado['Status'] ."</td>";
            $tabela .= '</tr>';
        }
        $tabela .= '</table>';

        //Grava o arquivo e dá permissão
        $file = "../../Docs/Excel/{$arquivo}";
        $fp = fopen($arquivo, 'w');
        $gravar = fwrite($fp, iconv("windows-1252","UTF-8",$tabela));
        fclose($fp);
        chmod($file, 0777);

        if($gravar){
            return true;
        }

        return false;


    }
}


