🧠 Prompt técnico do problema
Estou desenvolvendo um app em Streamlit que realiza o batimento entre fundos do CadFi e do Controle Espelho. Após identificar os fundos presentes em ambos, quero enriquecer esse relatório com os dados de protocolo e competência extraídos de um arquivo de balancete (XLSX ou PDF).

Já tenho uma função chamada enriquecer_em_comum_com_balancete que:

Normaliza os CNPJs dos dois DataFrames.
Remove duplicatas no balancete, mantendo o protocolo mais recente por fundo.
Faz o merge com base no CNPJ normalizado.
Preenche valores ausentes com "Não possui".
Posiciona as colunas Balancete_Protocolo e Balancete_Competencia após "Mes de Referencia".
O problema é que, mesmo com essa função implementada corretamente, o enriquecimento não está funcionando como esperado. Os protocolos e competências não estão sendo preenchidos no relatório final.

🧪 Suspeitas e hipóteses
Pode haver erro na normalização dos CNPJs (formato divergente entre os arquivos).
O balancete pode estar com colunas mal formatadas ou ausentes.
O merge pode estar sendo feito com CNPJs que não batem.
A função está sendo chamada incorretamente (ex: tentativa de importação de si mesma via from app import ...).
✅ Objetivo
Corrigir o fluxo para que o relatório "Em Ambos" seja enriquecido corretamente com os dados de protocolo e competência do balancete, por CNPJ.
