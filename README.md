Script criado em Python e interface utilizando TKINTER.

Programa feito para a empresa Barone Assessoria e Projetos (BAP), afim de calcular e gerar um documento de Memorial de Calculos Basicos para OODC e EIV/RIT-TIPO I-Lei 9.924/16 Prefeitura de Santo André.

	ESCLARECIMENTO PARA APLICAÇÃO E UTILIZADO DO APP. – “doc_BF_EIV”

Em Santo André com a criação da lei nº 8.696/2004 (plano Diretor) em seu artigo 122 passou a criar a condição para a construção além do Coeficiente de Aproveitamento Básico para as: Zonas de Qualificação Urbana; Zona de Recuperação Urbana e Zona de Reestruturação Urbana, chamada de Contrapartida financeira, correspondente a Outorga Onerosa do Direito de Construir, em 2.016 com a nova legislação através da Lei Ordinária nº 9.924/2016 – Lei de Uso, Ocupação e Parcelamento do Solo no Município de Santo André, altera e referencia o calculo para o Beneficio Financeiro estipulado para a tipologia de construção multifamiliar Vertical e também para os casos de construção verticais não residenciais.
A formula básica para o cálculo é:

BF = At x Vr x Cp x Ic x Fr , onde:

BF = 	Benefício Financeiro

At = 	Área do terreno

Vr = 	Valor de referência do metro quadrado do terreno para a aplicação da Outorga Onerosa do Direito de Construir

Cp	= 	Diferença entre o Coeficiente de Aproveitamento Pretendido e o Coeficiente de Aproveitamento Básico

Ic = Índice de Conversão de 0,4 (quatro décimos) para as Zonas de: Qualificação e Recuperação Urbana e de 0,33 para a Zona Reestruturação Urbana

Fr = Fator de Redução de 0,8 



Na Mesma Lei 9.924/2016 a partir do seu Art.º 38 ao Art.º 44 – “Capitulo III” – determina o que é EIV (Estudo de Impacto de Vizinhança) e qual são as aplicações e as tipologias de construção que estes estudos são aplicáveis e também a forma de apresentação, dentre eles temos o EIV tipo I. 
O EIV – Tipo I é um Estudo simplificado elaborado por meio do preenchimento de formulário a disponibilizado pela prefeitura através do link:
https://www.santoandre.sp.gov.br/portalservico/Alvara/DesenvolvimentoProjetosUrbanos/Requerimento_EIVX0911.aspx?TipoEmpr=N,  isento do atendimento dos Artºs. 40 a 42, cuja aprovação é condicionada ao atendimento da medida no percentual estabelecido no inciso XVIII e no § 3º do art. 43; 
Neste App Adicionamos o calculo do percentual adotado para as construções enquadradas no EIV tipo I. 









	ENTRADA DE DADOS E RESULTADO OBTIDO:

	ENTRADA DE DADOS:

Nome do projeto	-> 	nome adotado para vincular o projeto ao cálculo.

Área do terreno 	-> 	Área em m² do terreno em que o empreendimento será implantado.

Área Computável	 -> 	Área do empreendimento desconsiderando áreas comuns, áreas de circulação e subsolos utilizados como garagem.

Valor do FMP 	-> 	Valor fornecido pela Prefeitura de Santo André, reajustado sempre no primeiro dia do ano. O ano base do App é 2.024

Valor de referência  -> 	Valor Expresso em FMP Obtidos no Anexo 1.2 – MAPA 2 – VALOR DE REFERNCIA PARA CALCULO DE OUTROGA ONERODA DO DIREITO DE CONSTRUIR.  
Link p/ Anexo 1.2		: 	http://www4.cmsandre.sp.gov.br:9000/arquivo/31490 

Zona (1 ou2)  ->		(1) 	para imóveis situados em Zona de Qualificação Urbana e/ou Zona de Recuperação Urbana.
		(2) 	para imóveis situados em Zona de Reestruturação Urbana.
	
Area a Construir 	-> 	Area total do empreendimento.

Quanto ao índice CUB/SINDUSCON-SP o valor pode ser alterado mês a mês de acordo com o boletim econômico obtido no site do SINDUSCON-SP, neste aplicativo o valor e o padrão da construção podem ser alterados tendo como base o mês de para preenchimento do Formulário para EIV- TIPO I e que pode variar em função do termino da construção.


	RESULTADO OBTIDO:

Após a entrada de dados e clicar no botão Gerar documento, será criado um arquivo .docx para impressão e os resultados serão o valor do Benefício Financeiro para os índices de projeto e o Valor do EIV para o projeto pretendido, lembrando sempre que os valores finais vão sempre depender do prazo de execução da obra e de protocolo do processo junto a Prefeitura de Santo André.
