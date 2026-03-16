/* Script de extração de dados para sistema de ótica.
   Este código deve ser utilizado como um Bookmarklet no navegador.
*/

(function() {
    const getVal = (id) => document.getElementById(id)?.value || '';
    
    // Captura o nome do paciente/empresa
    let pacienteRaw = '';
    document.querySelectorAll('input').forEach(i => {
        if (i.value.includes('/') && i.readOnly && i.value.length > 10) {
            pacienteRaw = i.value;
        }
    });

    // Captura o ID da Ordem de Serviço (OS)
    let osId = '';
    document.querySelectorAll('.col-12.col-md-auto.text-truncate').forEach(div => {
        if (div.innerText.includes('ID do Pedido')) {
            osId = div.querySelector('.MediumTxt.HeaderColor')?.innerText || '';
        }
    });

    // Captura o nome da Lente e o Tratamento/Cor
    const lenteDesc = (document.querySelector('.MediumTxt.TextColor')?.innerText || '')
                      .replace(/Tipo da lente:/i, '').trim();
    const corTratamento = document.querySelector('label[for="cor"]')?.innerText || '';

    // Monta o objeto de dados
    const payload = {
        paciente: pacienteRaw,
        os: osId.replace(/\D/g, '').trim(),
        lente: lenteDesc,
        tratamento: corTratamento.trim(),
        esf_od: getVal('longeEsfericoOD'),
        cil_od: getVal('longeCilindricoOD'),
        eixo_od: getVal('longeEixoOD'),
        add_od: getVal('adicaoOD'),
        dnp_od: getVal('dnpOD'),
        alt_od: getVal('alturaOD'),
        esf_oe: getVal('longeEsfericoOE'),
        cil_oe: getVal('longeCilindricoOE'),
        eixo_oe: getVal('longeEixoOE'),
        add_oe: getVal('adicaoOE'),
        dnp_oe: getVal('dnpOE'),
        alt_oe: getVal('alturaOE')
    };

    // Envia para o servidor local Python (Flask)
    fetch('http://localhost:5000/salvar', {
        method: 'POST',
        mode: 'no-cors',
        body: JSON.stringify(payload)
    })
    .then(() => alert('✅ Dados da OS ' + payload.os + ' enviados com sucesso!'))
    .catch(() => alert('❌ Erro: Verifique se o servidor Python está ligado.'));
})();