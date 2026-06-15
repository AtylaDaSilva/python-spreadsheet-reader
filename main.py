

# ===========================================================================================
#                                      Leitor Planilha
# ===========================================================================================




# ===========================================================================================
#                                        Helpers / Outras classes
# ===========================================================================================


class EstilizadorDeCelula:
    def __init__(self, celula: Cell | MergedCell) -> None:
        self.celula = celula

    def alinhar(self, horizontal: str, vertical: str, quebrar_texo: bool):
        self.celula.alignment = Alignment(
            horizontal=horizontal, vertical=vertical, wrapText=quebrar_texo
        )

    def formatar(self, formato: str) -> None:
        self.celula.number_format = formato

