# desafio-controle-investimentos

# 📊 Simulador de Investimentos em Fundos Imobiliários (FIIs)

Ferramenta desenvolvida no Excel para simulação e projeção de investimentos em Fundos Imobiliários, desde o aporte mensal e acompanhamento de cenários de longo prazo até a distribuição da carteira conforme o perfil do investidor.

Este projeto foi desenvolvido como parte do Desafio de Projeto da **DIO (Digital Innovation One)** com o Expert Felipe Aguiar.

---

## 🎯 Perguntas de Negócio que a Ferramenta Responde

A planilha foi projetada para funcionar como uma aplicação simples e intuitiva, respondendo diretamente às seguintes questões:

1. **Quanto investir por mês?**  
   * Exibido no campo de entrada **Qual valor investir por mês?** (célula `D22` / intervalo `aporte`).
2. **Por quantos anos?**  
   * Definido no campo **Por quantos anos?** (célula `D23` / intervalo `qtd_anos`).
3. **Qual a taxa de rendimento mensal?**  
   * Informada no campo **Taxa de rendimento mensal** (célula `D24` / intervalo `taxa_mensal`).
4. **Quanto de patrimônio vai acumular?**  
   * Calculado automaticamente no campo **Patrimônio a acumular** (célula `D25` / intervalo `patrimonio`).
5. **Quanto vai receber de dividendos por mês?**  
   * Apresentado no campo **Dividendos mensais** (célula `D26`).

---

## ⚙️ Funcionalidades e Fórmulas Utilizadas

### 1. Cálculo de Patrimônio Acumulado (`VF` / `FV`)
Para projetar o valor futuro do montante acumulado ao final do período, utiliza-se a função financeira `VF` (Valor Futuro):
```excel
=VF(taxa_mensal; qtd_anos * 12; aporte * -1)

taxa_mensal: Taxa de juros aplicada a cada período mensal.qtd_anos * 12: Duração total da aplicação convertida em meses.aporte * -1: O valor do pagamento mensal (multiplicado por -1 para indicar a saída de caixa e retornar um resultado positivo).2. Cruzamento de Dados por Perfil (PROCV com Chave Composta)Para dividir o aporte entre os 6 tipos de Fundos Imobiliários (Papel, Tijolo, Híbrido, FOFs, Desenvolvimento e Hotelarias), foi criada uma aba de apoio (Tab_apoio_chave_composta).Cada linha da tabela de apoio gera uma chave composta única unindo o perfil e o tipo de FII:Excel=$B3 & "-" & $C3  -->  Exemplo: "Conservador-PAPEL"
Na aba principal (App), a percentagem sugerida é recuperada com o PROCV (ou VLOOKUP):Excel=PROCV($C$37 & "-" & B42; Tab_apoio_chave_composta!$A:$D; 4; FALSO)
Isso permite que, ao alterar o perfil no menu suspenso, toda a distribuição da carteira seja atualizada automaticamente.🏷️ Intervalos Nomeados CriadosPara tornar as fórmulas legíveis e amigáveis, foram configurados os seguintes intervalos nomeados na planilha:Nome do IntervaloCélula de OrigemDescriçãosalarioApp!$D$17Salário de referência para o planejamentorendimento_carteiraApp!$D$18Rendimento médio estimado da carteirasugestao_investimentoApp!$D$19Sugestão automática de aporte ($30\%$ do salário)aporteApp!$D$22Valor efetivo escolhido para investir mensalmenteqtd_anosApp!$D$23Período do investimento em anostaxa_mensalApp!$D$24Taxa de retorno mensal estimadapatrimonioApp!$D$25Patrimônio final acumulado calculado pela fórmula VF👤 Perfis de Investidor e Distribuição de AporteA soma dos percentuais de cada perfil resulta sempre em 100%:Tipo de FIIConservadorModeradoAgressivoPapel25%32%55%Tijolo50%30%10%Híbrido15%8%5%FOFs6%15%5%Desenvolvimento4%10%15%Hotelarias0%5%10%Total100%100%100%📈 Projeção de Cenários de Longo PrazoA ferramenta projeta a evolução do patrimônio e dos dividendos mensais em 5 janelas temporais distintas:2 Anos5 Anos10 Anos20 Anos30 Anos🖼️ Demonstração PráticaPerfil ConservadorSubstitua o caminho da imagem abaixo pelo link da sua captura de ecrã.Perfil AgressivoSubstitua o caminho da imagem abaixo pelo link da sua captura de ecrã.🚀 Como UtilizarFaça o download da planilha Desafio DIO - Controle de Investimentos.xlsx.Abra no Microsoft Excel ou Google Planilhas.Preencha apenas os campos das Configurações e do Investimento Mensal.Selecione o seu Perfil de Investidor no menu suspenso para visualizar a divisão sugerida do aporte.📝 LicençaEste projeto foi desenvolvido para fins educacionais durante o bootcamp da DIO.
