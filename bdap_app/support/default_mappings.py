"""Mappature predefinite lato Python per sorgenti dei template.

L'ordine delle chiavi riflette esattamente l'ordine di comparsa delle voci nel file
Template_Analisi.xlsx (foglio 'ANALISI CONTABILE' e foglio 'MAPPATURE_BDAP'),
seguito dalle voci del questionario (foglio 'VERIFICA QUESTIONARI' e dettagli).
"""

TEMPLATE_SOURCES: dict = {
    # -------------------------------------------------------------------------
    # 1. fondocassaal11 - Row 4 AC / Row 2 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fondocassaal11': {
        'cell_ref': 'D12',
        'expected_label': 'Fondo cassa al 1 gennaio',
        'label_check_cells': [
            'A12',
        ],
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 2. riscossioni - Row 5 AC / Row 3 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'riscossioni': {
        'cell_ref': 'D13',
        'expected_label': 'Riscossioni',
        'label_check_cells': [
            'A13',
        ],
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 3. pagamenti - Row 6 AC / Row 4 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'pagamenti': {
        'cell_ref': 'D14',
        'expected_label': 'Pagamenti',
        'label_check_cells': [
            'A14',
        ],
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 4. fondocassaal3112 - Row 7 AC / Row 5 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fondocassaal3112': {
        'cell_ref': 'D17',
        'expected_label': 'Fondo di cassa al 31 dicembre',
        'label_check_cells': [
            'A17',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 5. residuiattivi - Row 8 AC / Row 6 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuiattivi': {
        'cell_ref': 'D18',
        'expected_label': 'Residui attivi',
        'label_check_cells': [
            'A18',
        ],
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 6. residuiattivititietiiiep - Row 9 AC / Row 7 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuiattivititietiiiep': {
        'cell_ref': 'I5',
        'expected_label': 'TOTALE RESIDUI EP',
        'optional': True,
        'sheet_name_contains': 'FCDE',
        'source_workbook': 'analysis',
    },

    # -------------------------------------------------------------------------
    # 7. residuipassivi - Row 10 AC / Row 8 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuipassivi': {
        'cell_ref': 'D20',
        'expected_label': 'Residui passivi',
        'label_check_cells': [
            'A20',
        ],
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 8. fpvpartecorrenteal3112 - Row 11 AC / Row 9 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fpvpartecorrenteal3112': {
        'cell_ref': 'D21',
        'expected_label': 'Fondo pluriennale vincolato per spese correnti',
        'label_check_cells': [
            'A21',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'strict_expected_label': True,
    },

    # -------------------------------------------------------------------------
    # 9. fpvpartecapitaleal3112 - Row 12 AC / Row 10 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fpvpartecapitaleal3112': {
        'cell_ref': 'D22',
        'expected_label': 'Fondo pluriennale vincolato per spese in conto capitale',
        'label_check_cells': [
            'A22',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'strict_expected_label': True,
    },

    # -------------------------------------------------------------------------
    # 10. risultatodamministrazione - Row 13 AC / Row 11 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'risultatodamministrazione': {
        'cell_ref': 'D24',
        'expected_label': 'Risultato di amministrazione al 31 dicembre',
        'label_check_cells': [
            'A24',
        ],
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 11. prospettoevoluzionerisultatoammne - Row 14 AC / Row 12 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'prospettoevoluzionerisultatoammne': {
        'cell_ref': 'D24',
        'expected_label': 'Risultato di amministrazione al 31 dicembre',
        'label_check_cells': [
            'A24',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 12. parteaccantonatab - Row 15 AC / Row 13 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'parteaccantonatab': {
        'cell_ref': 'B33',
        'expected_label': 'Totale parte accantonata',
        'fallback_cell_refs': [
            'B36',
        ],
        'label_check_cells': [
            'A33',
            'A36',
        ],
        'optional': True,
        'row_tolerance': 6,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': True,
    },

    # -------------------------------------------------------------------------
    # 13. dicuifcde - Row 16 AC / Row 14 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'dicuifcde': {
        'cell_ref': 'B27',
        'expected_label': 'Fondo crediti di dubbia esigibilità al 31/12',
        'fallback_cell_refs': [
            'B28',
            'B29',
        ],
        'label_check_cells': [
            'A27',
            'A28',
            'A29',
        ],
        'optional': True,
        'row_tolerance': 4,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 14. partevincolatac - Row 17 AC / Row 15 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'partevincolatac': {
        'cell_ref': 'B40',
        'expected_label': 'Totale parte vincolata',
        'fallback_cell_refs': [
            'B43',
        ],
        'label_check_cells': [
            'A40',
            'A43',
        ],
        'optional': True,
        'row_tolerance': 6,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': True,
    },

    # -------------------------------------------------------------------------
    # 15. partedestinataagliinvestimentid - Row 18 AC / Row 16 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'partedestinataagliinvestimentid': {
        'cell_ref': 'B42',
        'default_if_missing': 0,
        'expected_label': 'Totale parte destinata agli investimenti',
        'fallback_cell_refs': [
            'B45',
        ],
        'label_check_cells': [
            'A42',
            'A45',
        ],
        'optional': True,
        'row_tolerance': 6,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': True,
    },

    # -------------------------------------------------------------------------
    # 16. partedisponibileeabcd - Row 19 AC / Row 17 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'partedisponibileeabcd': {
        'cell_ref': 'B43',
        'compute_if_missing': {
            'base': 'risultatoamministrazionea',
            'subtract': [
                'parteaccantonatab',
                'partevincolatac',
                'partedestinataagliinvestimentid',
            ],
        },
        'expected_label': 'Totale parte disponibile',
        'fallback_cell_refs': [
            'B46',
        ],
        'label_check_cells': [
            'A43',
            'A46',
        ],
        'optional': True,
        'row_tolerance': 6,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': True,
    },

    # -------------------------------------------------------------------------
    # 17. risultatodicompetenzaw1 - Row 20 AC / Row 18 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'risultatodicompetenzaw1': {
        'cell_ref': 'B69',
        'expected_label': 'W1) Risultato di competenza',
        'fallback_cell_refs': [
            'B76',
        ],
        'label_check_cells': [
            'A69',
            'A76',
        ],
        'optional': True,
        'row_tolerance': 10,
        'search_key_prefix': 'W1)',
        'sheet_idx': '7',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 18. equilibriodibilanciow2 - Row 21 AC / Row 19 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'equilibriodibilanciow2': {
        'cell_ref': 'B72',
        'expected_label': 'W2) Equilibrio di bilancio',
        'fallback_cell_refs': [
            'B81',
        ],
        'label_check_cells': [
            'A72',
            'A81',
        ],
        'optional': True,
        'row_tolerance': 10,
        'search_key_prefix': 'W2)',
        'sheet_idx': '7',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 19. equilibriocomplessivow3 - Row 22 AC / Row 20 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'equilibriocomplessivow3': {
        'cell_ref': 'B74',
        'expected_label': 'W3) Equilibrio complessivo',
        'fallback_cell_refs': [
            'B83',
        ],
        'label_check_cells': [
            'A74',
            'A83',
        ],
        'optional': True,
        'row_tolerance': 10,
        'search_key_prefix': 'W3)',
        'sheet_idx': '7',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 20. importocassavincolata - Row 23 AC / Row 21 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'importocassavincolata': {
        'cell_ref': 'D8',
        'expected_label': 'di cui cassa vincolata',
        'label_check_cells': [
            'A8',
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 4,
        'sheet_name_contains': [
            'SEZ. II-GEST. FIN. CASSA',
            'SEZ. II -DATI CASSA',
            'SEZ. II - Dati Cassa',
            'SEZ. II - GEST_FIN. 1',
        ],
        'source_workbook': 'questionario',
        'strict_expected_label': False,
        'table_refs': [
            {
            'cell_ref': 'D8',
            'year_cell_ref': 'D6',
        },
            {
            'cell_ref': 'D12',
            'year_cell_ref': 'D10',
        },
            {
            'cell_ref': 'E64',
            'year_cell_ref': 'E62',
        },
            {
            'cell_ref': 'D63',
            'year_cell_ref': 'D61',
        },
        ],
        'year_cell_ref': 'C6',
    },

    # -------------------------------------------------------------------------
    # 21. debitifuoribilancio - Row 24 AC / Row 22 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'debitifuoribilancio': {
        'cell_ref': 'P19',
        'expected_label': '4. Sono stati riconosciuti debiti fuori bilancio nel',
        'label_check_cells': [
            'B19',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_name': 'Sezione preliminare',
        'source_workbook': 'questionario_debiti',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 22. fcdeinrapportoaresiduiattivi - Row 25 AC / Row 23 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fcdeinrapportoaresiduiattivi': {
        'cell_ref': 'I6',
        'expected_label': [
            '% rispetto a residui attivi conservati (minimo 30%)',
            '% rispetto a residui attivi conservati',
            'FCDE IN RAPPORTO A RESIDUI ATTIVI',
        ],
        'is_percentage': False,
        'label_check_cells': [
            'A6',
            'B6',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_name_contains': 'FCDE',
        'source_workbook': 'analysis',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 23. fondocontenzioso - Row 26 AC / Row 24 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fondocontenzioso': {
        'cell_ref': 'I16',
        'default_if_missing': 0,
        'expected_label': [
            'Totale Fondo contenzioso',
            'Fondo contenzioso',
        ],
        'fallback_cell_refs': [
            'I14',
        ],
        'label_check_cells': [
            'B16',
            'B14',
        ],
        'only_fill_if_found': True,
        'optional': True,
        'row_tolerance': 10,
        'sheet_idx': '12',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 24. potenzialipassivita - Row 27 AC / Row 25 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'potenzialipassivita': {
        'cell_ref': 'I26',
        'default_if_missing': 0,
        'expected_label': [
            'FONDO ACCANTONAMENTO PASSIVITA POTENZIALI',
        ],
        'label_check_cells': [
            'D26',
        ],
        'only_fill_if_found': True,
        'optional': True,
        'row_tolerance': 10,
        'sheet_idx': '12',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': True,
    },

    # -------------------------------------------------------------------------
    # 25. fondogaranziadebiticommerciali - Row 28 AC / Row 26 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fondogaranziadebiticommerciali': {
        'cell_ref': 'I23',
        'default_if_missing': 0,
        'expected_label': [
            'Totale Fondo di garanzia debiti commerciali',
            'FONDO GARANZIA DEBITI COMMERCIALI',
            'Fondo di garanzia debiti commerciali',
        ],
        'fallback_cell_refs': [
            'I20',
            'I19',
        ],
        'label_check_cells': [
            'B23',
            'B20',
            'D19',
            'D23',
        ],
        'only_fill_if_found': True,
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '12',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 26. fondivincolaticovid - Row 29 AC / Row 27 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fondivincolaticovid': {
        'default_if_missing': 0,
        'expected_label': [
            'COVID',
            'FONDI COVID',
            'CONTRIBUTO STATO COVID',
            'SOLIDARIETA ALIMENTARE',
            'TRASFERIMENTI CORRENTI DA PRESIDENZA DEL CONSIGLIO DEI MINISTI -ORDINANZA PROTEZIONE CIVILE N. 658 DEL 29.03.2020 -EMERGENZA SANITARIA LEGATA AL DIFFONDERSI DEL CIRUS COVID-19 - MISURE URGENTI DI SOLIDARIETA ALIMENTARE',
        ],
        'optional': True,
        'row_tolerance': 10,
        'sheet_idx': '13',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
        'sum_cells': [
            'O',
        ],
    },

    # -------------------------------------------------------------------------
    # 27. fondocassadaidatisiope - Row 30 AC / Row 28 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fondocassadaidatisiope': {
        'default_if_missing': 'n.a.',
        'optional': True,
        'row_tolerance': 2.0,
        'source_workbook': 'rendiconto',
    },

    # -------------------------------------------------------------------------
    # 28. anticipazionetesoreria - Row 31 AC / Row 29 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'anticipazionetesoreria': {
        'cell_refs': [
            'B31',
        ],
        'expected_label': 'Titolo 7: Anticipazioni da istituto tesoriere/cassiere',
        'label_check_cells': [
            'A31',
        ],
        'optional': True,
        'row_tolerance': 10,
        'sheet_idx': '6',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 29. fal - Row 32 AC / Row 30 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'fal': {
        'cell_ref': 'B29',
        'expected_label': 'Fondo anticipazioni liquidità',
        'label_check_cells': [
            'A29',
        ],
        'optional': True,
        'row_tolerance': 10,
        'sheet_idx': '11',
        'source_workbook': 'rendiconto',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 30. contrazionenuovimutui - Row 33 AC / Row 31 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'contrazionenuovimutui': {
        'cell_ref': 'E25',
        'check_year_in_label': True,
        'expected_label': '3) Debito complessivo contratto',
        'fallback_cell_refs': [
            'F25',
        ],
        'label_check_cells': [
            'A25',
            'A25:C25',
        ],
        'optional': True,
        'row_tolerance': 10,
        'sheet_name_contains': [
            'SEZ. IV -INDEBITAMENTO_DATI _',
            'SEZ. IV -INDEBITAMENTO_DATI',
        ],
        'source_workbook': 'questionario',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 31. accantonamentocifrarelativaairinnovicontrattualidelccnl - Row 34 AC / Row 32 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'accantonamentocifrarelativaairinnovicontrattualidelccnl': {
        'cell_ref': 'I26',
        'default_if_missing': 0,
        'expected_label': [
            'FONDO RINNOVI CONTRATTUALI',
            'RINNOVI CONTRATTUALI',
            'ARRETRATI PER ANNI PRECEDENTI CORRISPOSTI AL PERSONALE A TEMPO INDETERMINATO',
            'ARRETRATI PER ANNI PRECEDENTI CORRISPOSTI AL PERSONALE',
        ],
        'label_check_cells': [
            'D26',
            'B26',
        ],
        'only_fill_if_found': True,
        'optional': True,
        'row_tolerance': 10,
        'sheet_idx': '12',
        'source_workbook': 'rendiconto',
        'static_column': True,
        'strict_expected_label': True,
    },

    # -------------------------------------------------------------------------
    # 32. spesadelpersonalesostenuta - Row 35 AC / Row 33 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'spesadelpersonalesostenuta': {
        'optional': True,
        'row_tolerance': 2.0,
        'source_workbook': 'relazione',
    },

    # -------------------------------------------------------------------------
    # 33. incidenzaspesapersonalesuspesacorrente - Row 36 AC / Row 34 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'incidenzaspesapersonalesuspesacorrente': {
        'as_decimal_ratio': True,
        'cell_ref': 'D27',
        'expected_label': 'Incidenza della spesa di personale sulla spesa corrente',
        'is_percentage': False,
        'label_check_cells': [
            'B27',
        ],
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '1',
        'source_workbook': 'indicatori',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 34. contrastoallevasionetributaria - Row 37 AC / Row 35 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'contrastoallevasionetributaria': {
        'cell_ref': 'D8',
        'divide_by_100': True,
        'expected_label': [
            'Percentuale di riscossione',
            'IMU/TASI recupero evasione',
            'Recupero evasione IMU/TASI',
            'TARSU/TIA/TARES/TARI recupero evasione',
            'Recupero evasione TARSU/TIA/TARES/TARI',
        ],
        'is_percentage': True,
        'label_check_cells': [
            'C8',
            'B6',
            'B12',
        ],
        'optional': True,
        'row_tolerance': 4,
        'sheet_name_contains': [
            'SEZ. II - DATI ENTRATE',
            'SEZ. II-DATI ENTRATE',
            'SEZ. II- DATI ENTRATE',
            'SEZ. II- GEST. ENTRATE',
        ],
        'source_workbook': 'questionario',
        'strict_expected_label': False,
        'table_refs': [
            {
                'cell_ref': 'D8',
                'year_cell_ref': 'D5',
            },
            {
                'cell_ref': 'D14',
                'year_cell_ref': 'D5',
            },
        ],
        'year_cell_ref': 'D5',
    },

    # -------------------------------------------------------------------------
    # 35. %diriscossionecomplessiva - Row 38 AC / Row 36 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    '%diriscossionecomplessiva': {
        'cell_ref': 'H58',
        'expected_label': 'Totale entrate',
        'is_percentage': True,
        'label_check_cells': [
            'B58',
        ],
        'optional': True,
        'row_tolerance': 5,
        'sheet_idx': '2',
        'sheet_name_contains': 'Entrate - Indicatori an',
        'source_workbook': 'indicatori',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 36. residuiattividatatiprecedenti2021_titolo_uno - Row 39 AC / Row 37 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuiattividatatiprecedenti2021_titolo_uno': {
        'tipo_tabella': 'attivi',
        'titolo': 'TITOLO I',
        'colonna_target': 'precedenti',
        'cell_ref': 'C71',
        'display_refs': [
            ('TITOLO I', 'C71'),
        ],
        'fallback_display_refs': [
            {
                'cell_ref': 'F6',
                'expected_label': 'Entrate Titolo 1.00',
                'label_check_cells': [
                    'B6',
                ],
                'row_tolerance': 2,
                'sheet_name_contains': [
                    'SEZ. II - Dati Cassa',
                    'SEZ. II -DATI CASSA',
                    'SEZ. II-DATI CASSA',
                ],
                'title': 'TITOLO I',
            },
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 2,
        'sheet_name_contains': [
            'SEZ. II -DATI RISULT. FINAN',
            'SEZ. II-DATI RISULT. FINAN',
            'SEZ. II - Dati Risult. Finan',
        ],
        'source_workbook': 'questionario',
        'strict_expected_label': False,
        'year_cell_ref': [
            'C69',
            'D69',
            'E69',
            'F69',
            'G69',
        ],
    },

    # -------------------------------------------------------------------------
    # 37. residuiattividatatiprecedenti2021_titolo_tre - Row 40 AC / Row 38 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuiattividatatiprecedenti2021_titolo_tre': {
        'tipo_tabella': 'attivi',
        'titolo': 'TITOLO III',
        'colonna_target': 'precedenti',
        'cell_ref': 'C73',
        'display_refs': [
            ('TITOLO III', 'C73'),
        ],
        'fallback_display_refs': [
            {
                'cell_ref': 'F10',
                'expected_label': 'Entrate Titolo 3.00',
                'label_check_cells': [
                    'B10',
                ],
                'row_tolerance': 2,
                'sheet_name_contains': [
                    'SEZ. II - Dati Cassa',
                    'SEZ. II -DATI CASSA',
                    'SEZ. II-DATI CASSA',
                ],
                'title': 'TITOLO III',
            },
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 2,
        'sheet_name_contains': [
            'SEZ. II -DATI RISULT. FINAN',
            'SEZ. II-DATI RISULT. FINAN',
            'SEZ. II - Dati Risult. Finan',
        ],
        'source_workbook': 'questionario',
        'strict_expected_label': False,
        'year_cell_ref': [
            'C69',
            'D69',
            'E69',
            'F69',
            'G69',
        ],
    },

    # -------------------------------------------------------------------------
    # 38. residuiattividatatiprecedenti2021_titolo_quattro - Row 41 AC / Row 39 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuiattividatatiprecedenti2021_titolo_quattro': {
        'tipo_tabella': 'attivi',
        'titolo': 'TITOLO IV',
        'colonna_target': 'precedenti',
        'cell_ref': 'C74',
        'display_refs': [
            ('TITOLO IV', 'C74'),
        ],
        'fallback_display_refs': [
            {
                'cell_ref': 'F12',
                'expected_label': 'Entrate Titolo 4.02.06',
                'label_check_cells': [
                    'B12',
                ],
                'row_tolerance': 2,
                'sheet_name_contains': [
                    'SEZ. II - Dati Cassa',
                    'SEZ. II -DATI CASSA',
                    'SEZ. II-DATI CASSA',
                ],
                'title': 'TITOLO IV',
            },
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 2,
        'sheet_name_contains': [
            'SEZ. II -DATI RISULT. FINAN',
            'SEZ. II-DATI RISULT. FINAN',
            'SEZ. II - Dati Risult. Finan',
        ],
        'source_workbook': 'questionario',
        'strict_expected_label': False,
        'year_cell_ref': [
            'C69',
            'D69',
            'E69',
            'F69',
            'G69',
        ],
    },

    # -------------------------------------------------------------------------
    # 39. residuipassividatati - Row 42 AC / Row 40 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuipassividatati': {
        'tipo_tabella': 'passivi',
        'titolo': 'TITOLO I',
        'colonna_target': 'precedenti',
        'cell_ref': 'C83',
        'display_refs': [
            ('TITOLO I', 'C83'),
        ],
        'fallback_display_refs': [
            {
                'cell_ref': 'F15',
                'expected_label': 'Spese Titolo 1.00 - Spese correnti',
                'label_check_cells': [
                    'B15',
                ],
                'row_tolerance': 2,
                'sheet_name_contains': [
                    'SEZ. II - Dati Cassa',
                    'SEZ. II -DATI CASSA',
                    'SEZ. II-DATI CASSA',
                ],
                'title': 'TITOLO I',
            },
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 4,
        'sheet_name_contains': [
            'SEZ. II -DATI RISULT. FINAN',
            'SEZ. II-DATI RISULT. FINAN',
            'SEZ. II - Dati Risult. Finan',
        ],
        'source_workbook': 'questionario',
        'strict_expected_label': False,
        'year_cell_ref': [
            'C81',
            'D81',
            'E81',
            'F81',
            'G81',
        ],
    },

    # -------------------------------------------------------------------------
    # Residui datati per il nuovo template ("Esercizi precedenti (n-4)")
    # -------------------------------------------------------------------------
    'residuiattividatatieserciziprecedentin4titoloi': {
        'tipo_tabella': 'attivi',
        'titolo': 'TITOLO I',
        'colonna_target': 'precedenti',
        'cell_ref': 'C63',
        'display_refs': [
            ('TITOLO I', 'C63'),
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 15,
        'sheet_name_contains': [
            'SEZ. II -DATI RISULT. FINAN',
            'SEZ. II-DATI RISULT. FINAN',
            'SEZ. II - Dati Risult. Finan',
        ],
        'source_workbook': 'questionario',
        'static_column': True,
        'strict_expected_label': False,
    },

    'residuiattividatatieserciziprecedentin4titoloiii': {
        'tipo_tabella': 'attivi',
        'titolo': 'TITOLO III',
        'colonna_target': 'precedenti',
        'cell_ref': 'C65',
        'display_refs': [
            ('TITOLO III', 'C65'),
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 15,
        'sheet_name_contains': [
            'SEZ. II -DATI RISULT. FINAN',
            'SEZ. II-DATI RISULT. FINAN',
            'SEZ. II - Dati Risult. Finan',
        ],
        'source_workbook': 'questionario',
        'static_column': True,
        'strict_expected_label': False,
    },

    'residuiattividatatieserciziprecedentin4titoloiv': {
        'tipo_tabella': 'attivi',
        'titolo': 'TITOLO IV',
        'colonna_target': 'precedenti',
        'cell_ref': 'C66',
        'display_refs': [
            ('TITOLO IV', 'C66'),
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 15,
        'sheet_name_contains': [
            'SEZ. II -DATI RISULT. FINAN',
            'SEZ. II-DATI RISULT. FINAN',
            'SEZ. II - Dati Risult. Finan',
        ],
        'source_workbook': 'questionario',
        'static_column': True,
        'strict_expected_label': False,
    },

    'residuipassividatatieserciziprecedentin4': {
        'tipo_tabella': 'passivi',
        'titolo': 'TITOLO I',
        'colonna_target': 'precedenti',
        'cell_ref': 'C75',
        'display_refs': [
            ('TITOLO I', 'C75'),
        ],
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 15,
        'sheet_name_contains': [
            'SEZ. II -DATI RISULT. FINAN',
            'SEZ. II-DATI RISULT. FINAN',
            'SEZ. II - Dati Risult. Finan',
        ],
        'source_workbook': 'questionario',
        'static_column': True,
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 40. ufficioavvocatura - Row 43 AC / Row 41 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'ufficioavvocatura': {
        'cell_ref': 'H67',
        'expected_label': 'L’ente dispone di un ufficio di avvocatura',
        'is_boolean_yes_no': True,
        'label_check_cells': [
            'B67',
        ],
        'default_if_missing': 'n.d.',
        'optional': True,
        'row_tolerance': 2,
        'sheet_name_contains': 'SPESA_PERSONALE',
        'source_workbook': 'questionario',
        'strict_expected_label': False,
    },

    # -------------------------------------------------------------------------
    # 41. incarichilegaliesterni - Row 44 AC / Row 42 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'incarichilegaliesterni': {
        'cell_ref': 'H73',
        'default_if_missing': 'n.d.',
        'expected_label': 'L’ente ha affidato incarichi',
        'is_boolean_yes_no': True,
        'label_check_cells': [
            'B73',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_name_contains': 'SPESA_PERSONALE',
        'source_workbook': 'questionario',
    },

    # -------------------------------------------------------------------------
    # 42. residuiattivifcde_titolo_uno - Row 45 AC / Row 43 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuiattivifcde_titolo_uno': {
        'cell_ref': 'D19',
        'expected_label': 'Totale Titolo 1: Entrate correnti di natura tributaria, contributiva e perequativa',
        'label_check_cells': [
            'B19',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_idx': '16',
        'source_workbook': 'rendiconto',
    },

    # -------------------------------------------------------------------------
    # 43. residuiattivifcde_titolo_tre - Row 46 AC / Row 44 MAPPATURE_BDAP
    # -------------------------------------------------------------------------
    'residuiattivifcde_titolo_tre': {
        'cell_ref': 'D35',
        'expected_label': 'Totale Titolo 3: Entrate extratributarie',
        'label_check_cells': [
            'B35',
        ],
        'optional': True,
        'row_tolerance': 2,
        'sheet_idx': '16',
        'source_workbook': 'rendiconto',
    },

    # =========================================================================
    # VOCI DEL QUESTIONARIO (foglio 'VERIFICA QUESTIONARI' e fallback)
    # =========================================================================
    '%riscossionisanzionidelcodicedellastrada': {
        'cell_ref': 'C20',
        'expected_label': 'Percentuale di riscossione',
        'is_percentage': True,
        'optional': True,
        'prefer_exact_table_workbook': True,
        'row_tolerance': 0,
        'sheet_name_contains': [
            'SEZ. II- DATI ENTRATE',
            'SEZ. II - DATI ENTRATE',
            'SEZ. II-DATI ENTRATE',
            'SEZ. II- Gest. ENTRATE',
            'SEZ. II - GEST. ENTRATE',
        ],
        'source_workbook': 'questionario',
        'strict_expected_label': False,
        'table_refs': [
            {
            'cell_ref': 'C20',
            'year_cell_ref': 'E5',
        },
            {
            'cell_ref': 'C24',
            'year_cell_ref': 'D15',
        },
            {
            'cell_ref': 'D25',
            'year_cell_ref': 'E16',
        },
            {
            'cell_ref': 'B32',
            'year_cell_ref': 'C23',
        },
        ],
        'year_cell_ref': 'E5',
    },

    'partenariatopubblicoprivato': {
        'source_workbook': 'questionario',
    },

    'disavanzodariaccertamentostraordinario': {
        'source_workbook': 'questionario',
    },

    'resacontoagenti': {
        'source_workbook': 'questionario',
    },

    'prospettocoperturaserviziadomanda': {
        'source_workbook': 'questionario',
    },

    'ufficioavvocatura': {
        'cell_ref': 'H67',
        'default_if_missing': 'n.d.',
        'expected_label': 'ufficio di avvocatura',
        'is_boolean_yes_no': True,
        'label_check_cells': ['B67', 'A67'],
        'optional': True,
        'row_tolerance': 5,
        'sheet_name_contains': ['SPESA_PERSONALE'],
        'source_workbook': 'questionario',
        'static_column': True,
        'strict_expected_label': False,
    },

    'incarichilegaliesterni': {
        'cell_ref': 'H73',
        'default_if_missing': 'n.d.',
        'expected_label': 'incarichi legali',
        'is_boolean_yes_no': True,
        'label_check_cells': ['B73', 'A73'],
        'optional': True,
        'row_tolerance': 5,
        'sheet_name_contains': ['SPESA_PERSONALE'],
        'source_workbook': 'questionario',
        'static_column': True,
        'strict_expected_label': False,
    },

}

# Alias per allineamento tra etichette template, MAPPATURE_BDAP e default_mappings
TEMPLATE_SOURCES['risultatoamministrazionea'] = TEMPLATE_SOURCES['prospettoevoluzionerisultatoammne']
TEMPLATE_SOURCES['fondocassaal3112f'] = TEMPLATE_SOURCES['fondocassaal3112']
TEMPLATE_SOURCES['prospettocassa'] = TEMPLATE_SOURCES['fondocassaal3112']
TEMPLATE_SOURCES['rispettoequilibri'] = TEMPLATE_SOURCES['risultatodicompetenzaw1']
TEMPLATE_SOURCES['residuiattivig'] = TEMPLATE_SOURCES['residuiattivi']
TEMPLATE_SOURCES['dicuiresiduiattivititieiiiepeserciziprecedenti'] = TEMPLATE_SOURCES['residuiattivititietiiiep']
TEMPLATE_SOURCES['residuipassivih'] = TEMPLATE_SOURCES['residuipassivi']
TEMPLATE_SOURCES['fpvpartecorrenteal3112i'] = TEMPLATE_SOURCES['fpvpartecorrenteal3112']
TEMPLATE_SOURCES['fpvpartecapitaleal3112l'] = TEMPLATE_SOURCES['fpvpartecapitaleal3112']
TEMPLATE_SOURCES['risultatodamministrazionemfghil'] = TEMPLATE_SOURCES['risultatodamministrazione']
TEMPLATE_SOURCES['fondopassivitpotenziali'] = TEMPLATE_SOURCES['potenzialipassivita']
TEMPLATE_SOURCES['fondocassadaidatisiope1450fondodicassadellenteallafinedelperiododiriferimentoquotavincolata'] = TEMPLATE_SOURCES['fondocassadaidatisiope']

# Nuove etichette per residui datati (Esercizi precedenti n-4)
TEMPLATE_SOURCES['residuiattividatatieserciziprecedentin4titoloi'] = TEMPLATE_SOURCES['residuiattividatatiprecedenti2021_titolo_uno']
TEMPLATE_SOURCES['residuiattividatatieserciziprecedentin4titoloiii'] = TEMPLATE_SOURCES['residuiattividatatiprecedenti2021_titolo_tre']
TEMPLATE_SOURCES['residuiattividatatieserciziprecedentin4titoloiv'] = TEMPLATE_SOURCES['residuiattividatatiprecedenti2021_titolo_quattro']
TEMPLATE_SOURCES['residuipassividatatieserciziprecedentin4'] = TEMPLATE_SOURCES['residuipassividatati']

# Alias per vecchie diciture del template
TEMPLATE_SOURCES['residuiattividatatiprecedentin4titoloi'] = TEMPLATE_SOURCES['residuiattividatatiprecedenti2021_titolo_uno']
TEMPLATE_SOURCES['residuiattividatatiprecedenti2021titolotree'] = TEMPLATE_SOURCES['residuiattividatatiprecedenti2021_titolo_tre']
TEMPLATE_SOURCES['residuiattividatatiprecedentin4titoloiv'] = TEMPLATE_SOURCES['residuiattividatatiprecedenti2021_titolo_quattro']

# Alias per riga 36 (indicatore 4.1 spesa del personale)
TEMPLATE_SOURCES['incidenzaspesapersonalesuspesacorrente40vindicatore41'] = TEMPLATE_SOURCES['incidenzaspesapersonalesuspesacorrente']
TEMPLATE_SOURCES['incidenzaspesapersonalesuspesacorrente40'] = TEMPLATE_SOURCES['incidenzaspesapersonalesuspesacorrente']

# Alias per righe 43-44
TEMPLATE_SOURCES['ufficioavvocaturaavvocatura'] = TEMPLATE_SOURCES['ufficioavvocatura']

