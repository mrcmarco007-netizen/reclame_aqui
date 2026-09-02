import streamlit.components.v1 as components


def leitor_cracha():

    components.html(
        """
        <script>

        const parentDocument = window.parent.document;


        function encontrarCampo() {

            const inputs =
                parentDocument.querySelectorAll('input');

            for (const input of inputs) {

                if (
                    input.type === 'text' &&
                    input.placeholder &&
                    input.placeholder.includes('crachá')
                ) {

                    return input;
                }
            }

            return null;
        }


        function encontrarBotao() {

            const botoes =
                parentDocument.querySelectorAll('button');

            for (const botao of botoes) {

                const texto =
                    (botao.innerText || '').trim();

                if (
                    texto.includes('REGISTRAR VOTO')
                ) {

                    return botao;
                }
            }

            return null;
        }


        function configurarLeitor() {

            const campo = encontrarCampo();

            if (!campo) {

                setTimeout(
                    configurarLeitor,
                    300
                );

                return;
            }


            campo.focus();


            if (
                campo.dataset.leitorConfigurado === 'true'
            ) {

                return;
            }


            campo.dataset.leitorConfigurado = 'true';


            campo.addEventListener(
                'keydown',
                function(event) {

                    if (event.key !== 'Enter') {

                        return;
                    }


                    event.preventDefault();


                    const valor =
                        campo.value.trim();


                    if (!valor) {

                        return;
                    }


                    /*
                     * Aguarda o Streamlit atualizar o valor
                     * do campo e então procura o botão.
                     */

                    setTimeout(
                        function() {

                            const botao =
                                encontrarBotao();


                            if (botao) {

                                botao.click();

                            }

                        },
                        300
                    );

                }
            );

        }


        configurarLeitor();


        /*
         * Mantém o campo preparado para o próximo crachá.
         */

        setInterval(
            function() {

                const campo =
                    encontrarCampo();

                if (campo) {

                    campo.focus();
                }

            },
            1000
        );

        </script>
        """,
        height=0
    )