# RateCheck
Este projeto é uma ferramenta de automação desenvolvida para otimizar o processo de auditoria e conciliação de tarifas hoteleiras. A aplicação cruza dados de relatórios internos (CSV) com confirmações de reserva acessadas via navegador web, garantindo que o valor cobrado corresponde ao valor confirmado.


# Hotel Rate Validator (Verificador de Tarifas)

Este projeto é uma ferramenta de automação desenvolvida para otimizar o processo de auditoria e conciliação de tarifas hoteleiras. A aplicação cruza dados de relatórios internos (CSV) com confirmações de reserva acessadas via navegador web, garantindo que o valor cobrado corresponde ao valor confirmado.

##  Funcionalidades Principais

* **Automação Web (RPA):** Utiliza **Selenium** para buscar automaticamente referências de reserva em um portal web/e-mail.
* **Extração Híbrida Inteligente:**
    * **Regex:** Identifica padrões de datas e valores monetários (BRL).
    * **Inteligência Artificial (NLP):** Integra a biblioteca **Transformers** (Hugging Face) como backup para interpretar textos complexos onde o Regex falha, respondendo a perguntas como "Qual a tarifa para a data X?".
* **Interface Gráfica Moderna (GUI):** Desenvolvida com **CustomTkinter**, oferecendo modo escuro/claro, abas de navegação e feedback visual em tempo real.
* **Processamento de Dados:** Leitura e tratamento de arquivos CSV com **Pandas**, incluindo lógica para ignorar quartos "Share" (múltiplos hóspedes) ou lista de exclusão manual.
* **Relatórios:** Gera um resumo visual (Treeview) com status coloridos (Correto, Erro de Tarifa, Sem Referência).

##  Tecnologias Utilizadas

* **Python 3.12+**
* **GUI:** `customtkinter`, `tkinter`, `ttkbootstrap`
* **Automação:** `selenium` (Webdriver Manager)
* **Dados:** `pandas`, `re` (Regex)
* **AI/ML:** `transformers`, `torch` (DistilBERT model)

##  Como Funciona

Navegue até Bookings > Reservations > Manage Reservation.

Configure os filtros de busca:

Arrival From: Coloque uma data de início abrangente (ex: 01/01/2025).

Arrival To: Coloque a data atual da conferência (ex: 05/12/2025).

Reservation Status: Selecione IN HOUSE.

Clique em Search.

Para exportar:

Vá em View Options > Export > CSV.

Selecione Loaded Rows e clique em Export.

 Atenção: O sistema carrega apenas 100 reservas por vez. É necessário rolar a página para carregar mais reservas e repetir a exportação para garantir que todos os dados sejam capturados.


1.  O usuário carrega os arquivos CSV contendo as reservas do dia.
2.  Define uma "Data Alvo" para a conferência.
3.  O script se conecta a uma sessão de navegador existente (via Debugger Address) para evitar bloqueios de login.
4.  Para cada reserva, o sistema:
    * Busca o localizador (External Reference).
    * Abre o detalhe da reserva/e-mail.
    * Lê o corpo do texto e extrai a tarifa diária correspondente à data alvo.
    * Compara com o valor presente no CSV.
5.  O resultado é exibido na tela e classificado automaticamente.

É necessário abrir uma aba do navegador no modo depuração com a porta para o selenium 

"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebug"


