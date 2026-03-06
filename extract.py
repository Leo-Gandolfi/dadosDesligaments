import fitz
import pandas as pd
import re
import sys

def extract_data(pdf_path, excel_out_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    # Split by the header to get individual records
    records_raw = re.split(r"Solicita..o de Desligamento - (\d+)", full_text)
    
    # records_raw[0] is empty or garbage before the first match
    # From index 1 onwards, they are in pairs: (SolicitacaoID, Text)
    
    data = []
    
    for i in range(1, len(records_raw), 2):
        sol_id = records_raw[i]
        rec_text = records_raw[i+1]
        
        # We need to find the block of 'keys' and 'values'
        # The values start right after "Haver. Reposi..o:"
        # There are 9 keys, so 9 values.
        
        lines = [line.strip() for line in rec_text.split('\n') if line.strip()]
        
        # Find index of "Haver. Reposi..o:"
        idx_havera = -1
        for j, line in enumerate(lines):
            if re.search(r"Haver.\s+Reposi..o:", line):
                idx_havera = j
                break
                
        if idx_havera != -1:
            try:
                # The values are the next 9 lines
                val_status = lines[idx_havera + 1]
                val_colaborador = lines[idx_havera + 2]
                val_cc = lines[idx_havera + 3]
                val_cargo = lines[idx_havera + 4]
                val_data_adm = lines[idx_havera + 5]
                val_tipo_deslig = lines[idx_havera + 6]
                val_motivo = lines[idx_havera + 7]
                val_liberacao = lines[idx_havera + 8]
                val_havera_rep = lines[idx_havera + 9]
                
                # NP and Colaborador
                np_colab_match = re.match(r"(\d+)\s*[-–]\s*(.*)", val_colaborador)
                if np_colab_match:
                    np_val = np_colab_match.group(1).lstrip("0")
                    colab_val = np_colab_match.group(2)
                else:
                    np_val = ""
                    colab_val = val_colaborador
                    
            except IndexError:
                continue
        else:
            continue
            
        # Extract Outras Informações
        # This is between "Recomendaria-o a outros setores?" block and "Administra..o de Pessoal"
        # First let's find "Administra..o de Pessoal"
        try:
            idx_admin = next(k for k, line in enumerate(lines) if re.search(r"Administra..o de Pessoal", line))
            # The Outras Informacoes value is right above idx_admin, sometimes can be multiple lines.
            # We can trace back until "Recomendaria-o a outros setores?" or "Sim" or "N.o" 
            # But usually it's just the line right before:
            val_outras_info = ""
            for k in range(idx_admin - 1, -1, -1):
                if lines[k] in ["Sim", "No", "Não", "No"] or re.search(r"Recomendaria-o", lines[k]):
                    break
                val_outras_info = lines[k] + " " + val_outras_info
            
            val_outras_info = val_outras_info.strip()
            if val_outras_info in ["Outras Informaes", "Outras Informações", "Outras Informaes"]:
                val_outras_info = "" # It was empty
        except StopIteration:
            val_outras_info = ""
            
        # Recrutamento e Seleção - Condições de Readmissão
        # Diretoria/Gerência - Condições de Readmissão
        # Considerações
        
        # Let's search for "Recrutamento e Sele..o" -> next is "Condi..es de Readmiss..o" -> next is value
        val_recrut_readm = ""
        for j, line in enumerate(lines):
            if re.search(r"Recrutamento e Sele..o", line):
                # Next line is usually "Condições de Readmissão", next is value (unless it reaches next section)
                if j + 2 < len(lines):
                    if lines[j+2] != "Diretoria /Gerncia de Unidade de Negcio" and "Diretoria" not in lines[j+2]:
                        val_recrut_readm = lines[j+2]
                break

        # Diretoria/Gerência - Considerações
        val_dir_readm = ""
        val_consid = ""
        aprovador_id = ""
        aprovador_nome = ""

        idx_dir = -1
        for j, line in enumerate(lines):
            if re.search(r"Diretoria\s*/Ger.ncia de Unidade de Neg.cio", line):
                idx_dir = j
                break

        if idx_dir != -1:
            # Find "Considerações" after the Diretoria header
            idx_consid = -1
            for j in range(idx_dir + 1, len(lines)):
                if re.search(r"Considera..es", lines[j]):
                    idx_consid = j
                    break

            # "Condições de Readmissão" label sits between idx_dir and idx_consid;
            # collect everything between that label and "Considerações" as the value.
            if idx_consid != -1:
                for j in range(idx_dir + 1, idx_consid):
                    if re.search(r"Condi..es de Readmiss.o", lines[j]):
                        readm_parts = [lines[k] for k in range(j + 1, idx_consid)]
                        val_dir_readm = " ".join(readm_parts).strip()
                        break

            # Collect ALL lines of Considerações until "Fluxo de Aprova" or end
            if idx_consid != -1:
                consid_parts = []
                for k in range(idx_consid + 1, len(lines)):
                    if re.search(r"Fluxo de Aprova", lines[k]):
                        break
                    consid_parts.append(lines[k])
                if consid_parts:
                    val_consid = consid_parts[0] + " - " + " ".join(consid_parts[1:]) if len(consid_parts) > 1 else consid_parts[0]
                else:
                    val_consid = ""

        # Aprovador and Nome do Aprovador: ALWAYS the first row in "Fluxo de Aprovações"
        # Strategy: find the Fluxo header, then scan for the FIRST date (DD.MM.YYYY).
        # The line immediately before the date is the Nome Aprovador.
        # If the line before that is a 6-digit number, it's the Aprovador ID.
        idx_fluxo = -1
        for j, line in enumerate(lines):
            if re.search(r"Fluxo de Aprova", line):
                idx_fluxo = j
                break

        if idx_fluxo != -1:
            scan_start = idx_fluxo + 1  # start right after the header
            for k in range(scan_start, len(lines)):
                if re.match(r"^\d{2}\.\d{2}\.\d{4}$", lines[k]):
                    # Found first date — the name is one line before
                    aprovador_nome = lines[k - 1]
                    # Check if line before name is a numeric ID
                    if k >= 2 and re.match(r"^\d{6,}$", lines[k - 2]):
                        aprovador_id = lines[k - 2].lstrip("0")
                    break

        data.append({
            'Solicitacao de Desligamento': sol_id,
            'NP': np_val,
            'Colaborador': colab_val,
            'Havera reposicao?': val_havera_rep,
            'Tipo de Desligamento': val_tipo_deslig,
            'Motivo': val_motivo,
            'Outras informacoes': val_outras_info,
            'Recrutamento e Selecao - Condicoes de Readmissao': val_recrut_readm,
            'Diretoria/Gerencia - Condicoes de Readmissao': val_dir_readm,
            'Consideracoes': val_consid,
            'Aprovador': aprovador_id,
            'Nome do Aprovador': aprovador_nome
        })
        
    df = pd.DataFrame(data)
    print(f"Extracted {len(df)} records.")
    
    # Save to Excel
    original_df = pd.read_excel('EXEMPLO DO RELATÓRIO 2.xlsx')
    cols = original_df.columns
    
    # Re-map the dict to use exact columns from the output file
    mapped_data = []
    for row in data:
        mapped_data.append({
            cols[0]: row['Solicitacao de Desligamento'],
            cols[1]: row['NP'],
            cols[2]: row['Colaborador'],
            cols[3]: row['Havera reposicao?'],
            cols[4]: row['Tipo de Desligamento'],
            cols[5]: row['Motivo'],
            cols[6]: row['Outras informacoes'],
            cols[7]: row['Recrutamento e Selecao - Condicoes de Readmissao'],
            cols[8]: row['Diretoria/Gerencia - Condicoes de Readmissao'],
            cols[9]: row['Consideracoes'],
            cols[10]: row['Aprovador'],
            cols[11]: row['Nome do Aprovador']
        })
        
    out_df = pd.DataFrame(mapped_data)
    
    # Append
    final_df = pd.concat([original_df, out_df], ignore_index=True)
    final_df.to_excel(excel_out_path, index=False)
    print(f"Saved to {excel_out_path}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = "TESTE- 35_merged.pdf"
        
    if len(sys.argv) > 2:
        out_path = sys.argv[2]
    else:
        out_path = "EXEMPLO DO RELATÓRIO 2_output.xlsx"
        
    extract_data(pdf_path, out_path)


