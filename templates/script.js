document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    const uploadBtn = document.getElementById('uploadBtn');
    const statusDiv = document.getElementById('status');
    const reportsList = document.getElementById('reportsList');

    // Função para carregar lista de relatórios
    function loadReports() {
        fetch('/relatorios_analiticos/')
            .then(response => response.text())
            .then(html => {
                const parser = new DOMParser();
                const doc = parser.parseFromString(html, 'text/html');
                const links = doc.querySelectorAll('a');
                const reportLinks = Array.from(links)
                    .filter(link => link.href.endsWith('.pdf'))
                    .map(link => ({
                        name: link.textContent,
                        url: link.href
                    }));

                if (reportLinks.length > 0) {
                    const list = document.createElement('ul');
                    list.className = 'reports-list';
                    reportLinks.forEach(report => {
                        const li = document.createElement('li');
                        const a = document.createElement('a');
                        a.href = report.url;
                        a.textContent = report.name;
                        a.target = '_blank';
                        li.appendChild(a);
                        list.appendChild(li);
                    });
                    reportsList.innerHTML = '';
                    reportsList.appendChild(list);
                } else {
                    reportsList.innerHTML = '<p>Nenhum relatório gerado ainda.</p>';
                }
            })
            .catch(() => {
                reportsList.innerHTML = '<p>Não foi possível carregar os relatórios.</p>';
            });
    }

    // Carregar relatórios ao iniciar
    loadReports();

    uploadBtn.addEventListener('click', function() {
        const files = fileInput.files;
        if (files.length === 0) {
            showStatus('Por favor, selecione pelo menos um arquivo.', 'error');
            return;
        }

        const formData = new FormData();
        for (let i = 0; i < files.length; i++) {
            formData.append('files[]', files[i]);
        }

        showStatus('Enviando arquivos...', 'info');

        // Aqui você faria o upload real para o GitHub
        // Por enquanto, simulamos o processo
        setTimeout(() => {
            showStatus(`Arquivos enviados com sucesso! Os relatórios serão gerados em breve.`, 'success');
            // Recarregar lista após alguns segundos
            setTimeout(loadReports, 5000);
        }, 2000);
    });

    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = type;
    }
});
