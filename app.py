
                # Fórmula
                if (
                    isinstance(
                        valor_original,
                        str,
                    )
                    and valor_original.startswith("=")
                ):
                    novo_valor = (
                        manter_formula_original(
                            valor_original
                        )
                    )

                else:
                    valor_decimal = (
                        converter_valor_excel_para_decimal(
                            valor_original
                        )
                    )

                    if valor_decimal is None:
                        novo_valor = (
                            "Não será alterado"
                        )
                    else:
                        novo_valor = (
                            arredondar_sempre_para_cima(
                                valor_decimal
                                * Decimal(
                                    str(multiplicador)
                                )
                            )
                        )

                dados_previa.append(
                    {
                        "Coluna": letra_coluna,
                        "Linha": numero_linha,
                        "Célula": (
                            letra_coluna
                            + str(numero_linha)
                        ),
                        "Valor original": valor_original,
                        "Novo valor": novo_valor,
                    }
                )

        if dados_previa:
            st.subheader("Prévia")

            st.dataframe(
                pd.DataFrame(dados_previa),
                use_container_width=True,
                hide_index=True,
            )

        else:
            st.warning(
                "Nenhum valor foi encontrado "
                "nas colunas selecionadas."
            )

        # ==================================================
        # GERAR EXCEL
        # ==================================================

        if st.button(
            "Gerar Excel atualizado",
            type="primary",
        ):

            with st.spinner(
                "Atualizando colunas "
                + ", ".join(
                    letras_colunas_selecionadas
                )
                + "..."
            ):

                resultado, conferencias = (
                    atualizar_excel(
                        arquivo_excel=conteudo_arquivo,
                        nome_aba=nome_aba,
                        linha_inicio=linha_inicio,
                        numeros_colunas=numeros_colunas,
                        multiplicador=multiplicador,
                    )
                )

            if not conferencias:
                st.warning(
                    "Nenhuma célula foi alterada. "
                    "Verifique as colunas e a linha inicial."
                )
                st.stop()

            st.success(
                "Excel pronto. "
                + str(len(conferencias))
                + " célula(s) processada(s)."
            )

            st.subheader("Conferência")

            st.dataframe(
                pd.DataFrame(conferencias),
                use_container_width=True,
                hide_index=True,
            )

            nome_saida = (
                Path(nome_arquivo).stem
                + "_precos_atualizados.xlsx"
            )

            st.download_button(
                label="⬇️ Baixar Excel atualizado",
                data=resultado,
                file_name=nome_saida,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )

    except Exception as erro:
