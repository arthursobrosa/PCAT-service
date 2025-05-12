function updateAcronyms(agentType, acronymSelect, acronymsConcessionaria, acronymsPermissionaria) {
    acronymSelect.innerHTML = '';

    if (!agentType) {
        acronymSelect.disabled = true;
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = 'Selecione o tipo de agente primeiro';
        acronymSelect.appendChild(defaultOption);
        return;
    }

    acronymSelect.disabled = false;

    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Selecione a empresa';
    acronymSelect.appendChild(defaultOption);

    let acronyms = [];

    if (agentType == "Concessionária") {
        acronyms = acronymsConcessionaria;
    } else if (agentType == "Permissionária") {
        acronyms = acronymsPermissionaria;
    }

    acronyms.forEach(function (acronym) {
        const option = document.createElement('option');
        option.value = acronym;
        option.textContent = acronym;
        acronymSelect.appendChild(option);
    });
}


document.addEventListener('DOMContentLoaded', function () {
    const agentTypeSelect = document.getElementById('agent_type');
    const acronymSelect = document.getElementById('acronym');

    acronymSelect.disabled = true;

    agentTypeSelect.addEventListener("change", function () {
        updateAcronyms(
            agentTypeSelect.value,
            acronymSelect,
            acronymsConcessionaria,
            acronymsPermissionaria
        );
    });

    const form = document.querySelector("form");
    const loadingIndicator = document.getElementById('loading');

    form.addEventListener("submit", function (e) {
        if (!agentTypeSelect.value) {
            alert("Por favor, selecione o tipo de agente.");
            e.preventDefault();
            return;
        }

        if (!acronymSelect.value) {
            alert("Por favor, selecione a empresa.");
            e.preventDefault();
            return;
        }

        loadingIndicator.textContent = "Processando..."
        loadingIndicator.style.display = 'block';
    });

    window.addEventListener('pageshow', function (event) {
        if (event.persisted && loadingIndicator.style.display === 'block') {
            loadingIndicator.style.display = 'none';
        }
    });
});