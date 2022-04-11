import pandas as pd
import numpy as np
import re

out='C:\\Users\\sina.kian\\Desktop\\Quick report format V2\\update'
C_Group=['Group ID', 'Group Name']
C_Company=['Company ID', 'Company Name', 'Company Code', 'Group ID', 'Complete Address ID', '% Part.', 'Tramite', 'Indrizzo SS', 'n.c. SS', 'cap SS', 'Site SS', 'pv SS', 'State SS', 'nazione SS', 'Longitude SS', 'Latitude SS', 'Accuracy SS', 'PAS', 'Area of Business', 'CS', 'CS Currency', 'P.IVA', 'Note', 'FT', 'FVT', 'FVT P Propri', 'FVT P meta-Propri', 'FVT P Terzi', 'FVT P prima materie terzi', 'FVT P altri', 'FVT Italy', 'FVT European Unioin (ex taly)', 'FVT USA Canada', 'FVT Rest of World', 'RU-Dirigenti', 'RU-Quadri', 'RU-Impiegati', 'RU-Operai', '(RAL) D €', '(RAL) Q €', '(RAL) I €', '(RAL) O €', 'CCNL D', 'CCNL Q', 'CCNL I', 'CCNL O', 'BE D', 'BE Q', 'BE I', 'BE O', 'Classificazione Inail', 'Tasso applicato', 'Tariffa', 'Anno di riferimento (INAIL)']
C_Locations=['Complete Address ID','Complete Address','City','Group ID','Company','Company ID','Address','n°','frazione località  specifiche','cap','pv','nazione','Longitude','Latitude','Accuracy','Construction Year','Building material','bene culturale','Pollution Classification','Ricavi infragruppo','N. Dipendenti','soggetto proprietario','Site type','Activity situation','Activity Type','Building Ownership','Third party use','Number of Floors','SST[m2]','SSS[m2]','DUG','DUS','Fabbricato','Contenuto','Merci','Total Value','Business interruption','Total including BI','Polizza','RM Max' ,'SF Max' ,'FP Max' ,'FPS Max' ,'TPP Max' ,'TOTAL products Max','Squadra interna emergenza','Permesso fumo','Piano emergenza','Esercitazione annuale','Pianificazione programmi manutenzione' ,'Spinkler','Idranti esterni','Manichette interne','Riserve acqua antincendio','esclusiva antincendio','elettriche','diesel','generatore emergenza','Rilevatori incendio/fumo','R.I.F Coefficient','R.I.F Category','Estintori','Sistemi speciali protezione','Cat SSP','Rilevatori perimetrali','Rilevatori microonde interni','Sistemi accesso controllato','Illuminazione su allarme','Sirena','Sirena Category','TVCC','Number of TVCC','TVCC Colori','TVCC Night Vision','registrazione immagini','Servizio vigilanza','V.F più vicina','distanza (km)','tempo intervento [min] ','dimensione riserva [m3] ','Destinatario allarme']
C_Transportation=['Transportation ID','Group ID','Company ID','Transaction Type','From Country','To Country','Destinario DF','Intermediate Country','IC Description','Fornitore','SP category','Mezzo Trasporto','Fr. Destino','EXW','FOB Ge/La S.','Accordo trasporto latte','Max transported value','Currency Trasported','Valore Volumi annui','unità misura VA','F/C Beni annuale','Currency F/CBA','Note Tras','Ricavi in 1E6 euro']
C_Prodotto=['Product ID','Tipologia di prodotto','Group ID','Active or Outdated','Cluster di prodotto','Linea Prodotto','Company ID','Produzione propria','Commercializzazione prodotti di terzi','Settori di impiego','valore Quantità vendute','Unit Quantità Vendute','R Italia','R EU','R remaining Europe','R USA  Canada','R Rest of the world','R Totale','Primo anno vendita','Vita Utile Media','Unit Vita Utile','Diretta al cliente','Grossisti','Commercializzata terzi marchio','Altro tipo vendita','Etichettatura','Confezionamento','Altro','third party deposit control','Prodotto indipendente','assemblatore produttore finale','Strumento controllo misurazione','Altro dipendenza','PF Terzi Componente','Destinazione uso PF Terzi','Incidenza valore PF Terzi','contenuta PF di terzi','Modalità incorporazione PF Terzi','Utilizzo da parte Terzo','Principali aziende terze utilizzatrici','Caratteristiche indicate','venduti con istruzioni','Commento','Fenomeno','Causa','Avvertenze',' soggetti controllo qualità','Marchio di qualità','PP Interna','PP Esterna','Standard Applicati','Certificazioni prodotto','Documentazione','Anni Documentazzione','Autorizatore Documenti','Confronto Concorrenza','Risk Accidential Use','Risk Overload Overheating','Risk Other Use','Risk dangerous environment use','Risk inherent product use','Comprehensible Instructions','Misuse Instructions','Translation Reviews','Periodic safety Functionality Reviews','procedure scritte','Periodic Recalibration','Product Samples','CMP A campione','CMP Per lotto','CMP In linea','CMP Mutipli combinati','CMP Altro','CS A campione','CS Per lotto','CS In linea','CS Mutipli combinati','CS Altro','CPF A campione','CPF Per lotto','CPF In linea','CPF Mutipli combinati','CPF Altro','CLP A campione','CLP Per lotto','CLP In linea','CLP Mutipli combinati','CLP Altro','CTVRU A campione','CTVRU Per lotto','CTVRU In linea','CTVRU Mutipli combinati','CTVRU Altro','Tipo Test Utilizzatore','CS Installazione','CS Manutenzione','CS Corsi clienti','CS Corsi subapp','CS Asistenza vendita','Ufficio Revisione','limitazione responsabilità vendita','limitazione responsabilità manuali','limitazione responsabilità brochure','CLR Da Fornitori','CLR Da Distributori','CLR Da Clienti','CLR A Fornitori','CLR A Distributori','CLR A Clienti','Tracciabilità prodotti','Grossista Dettagliante Utilizzatore','potenziali pericoli rilevatisi','Livello Sicurezza Aggiornate','Modifiche Sostituzione prodotti','DSP Prodotti Difettosi','DSP Difetto prodotto','DSP UI MI','DSP Difetto stoccaggio','DSP Altro','DSP ufficio reclami','DSP Reclami Sinistri avvenuti','DSP Reclami pendenti','DSP autorità ritiro prodotto','DSP Ritiro prodotto','DSP Date','IE Tipologia intervento autorità','IE Descrizione intervento','IE Autorità richiedente esecutrice','II Tipologia','Tipologia attività presso terzi','Fatturato attività presso terzi','Costo attività presso terzi','Attività da terzi','soggetti da terzi utilizzati','Luogo esecuzione da Terzi','Costo esercizio da Terzi','prodotti di terzi utilizzati','soggetti terzi fornitori','Costo esercizio PRODOTTI DI TERZI','FVT Poduct Type','FVT PT [company ID]']
C_Revenue_Country=['Group ID','Country','Continent','Fatturato Netto','Currency F.N','Self production sold quantity','Third party sold quantity','Total sold quantity','Measuring Unit','M Propri Quote','AMP Quote','MA Quote','M Privati Quote','MF Quote','M Propri Fatturato','AMP Fatturato','MA Fatturato','M Privati Fatturato','MF Fatturato','Fatturato Currency Marchio','M Propri Quantità','AMP Quantità','MA Quantità','M Privati Quantità','MF Quantità','Misura di quantità']
C_Dependenza_Clienti=['Group ID','N. Cliente','Fatt. Netto Anno x','Fatt. Netto Anno y','Incidenza fatt Anno y su tot','Delta Fatturato Valore Anno y-Anno x','Delta Fatturato % Anno y-Anno x','Q.tà Anno x','Q.tà Anno y','Delta Q.tà Valore Anno y-Anno x','Delta Q.tà % Anno y-Anno x']
C_Sinistri=['Group ID','Sinistro ID','Tipologia Generale','Società interessata','Ubicazione interessata','Anno di accadimento','Anno di denuncia','Causa','Effetto','Tipologia di evento','Ramo','Provenienza','Vettore','Prodotto','Terzo danneggiato','IDAN subito dall\'Azienda','IDAN richiesto dal terzo','IDAN riservato dall\'assicuratore','IDAN risarcito da assicurazione','IDAN indennizzato da assicurazione','IDAN Franchigia applicata','IDAN Currency','PI Numero','PI Tipo','PI Compagnia','Stato del sinistro','Descrizione sinistro','Procedimento giudiziario','Persone decedute']

Group = pd.DataFrame(columns=C_Group)
Company = pd.DataFrame(columns=C_Company)
Locations = pd.DataFrame(columns=C_Locations)
Transportation = pd.DataFrame(columns=C_Transportation)
Prodotto = pd.DataFrame(columns=C_Prodotto)
Revenue_Country = pd.DataFrame(columns=C_Revenue_Country)
Dependenza_Clienti = pd.DataFrame(columns=C_Dependenza_Clienti)
Sinistri = pd.DataFrame(columns=C_Sinistri)

writer = pd.ExcelWriter('Empty_DB.xlsx', engine='xlsxwriter')

Group.to_excel(writer, sheet_name='Group')
Company.to_excel(writer, sheet_name='Company')
Locations.to_excel(writer, sheet_name='Locations')
Transportation.to_excel(writer, sheet_name='Transportation')
Prodotto.to_excel(writer, sheet_name='Prodotto')
Revenue_Country.to_excel(writer, sheet_name='Revenue_Country')
Dependenza_Clienti.to_excel(writer, sheet_name='Dependenza_Clienti')
Sinistri.to_excel(writer, sheet_name='Sinistri')

df_map = {'Group': Group, 'Company': Company, 'Locations': Locations, 'Transportation': Transportation, 'Prodotto': Prodotto, 'Revenue_Country': Revenue_Country, 'Dependenza_Clienti': Dependenza_Clienti, 'Sinistri': Sinistri,}
for name, df in df_map.items():
    #df.to_pickle(os.path.join(out, f'\\{name}.pkl'))
    df.to_pickle(name+'.pkl')



writer.save()

#import Fill
