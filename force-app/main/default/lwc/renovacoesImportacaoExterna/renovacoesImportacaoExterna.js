import { LightningElement } from 'lwc';
import { loadScript } from 'lightning/platformResourceLoader';
import XLSX_RESOURCE from '@salesforce/resourceUrl/renovacoesXlsxLib';

const PREVIEW_LIMIT = 10;
const NONE_VALUE = '__NONE__';

const MONTH_NAMES = [
    'Jan',
    'Fev',
    'Mar',
    'Abr',
    'Mai',
    'Jun',
    'Jul',
    'Ago',
    'Set',
    'Out',
    'Nov',
    'Dez'
];


const PROCESS_FIELDS = [

    {
        apiName: 'numeroOrcamento',
        label: 'Nº orçamento',
        description: 'Número do orçamento ou proposta na seguradora.',
        required: true,
        aliases: [
            'orcamento',
            'numeroorcamento',
            'numerodoorcamento',
            'codigoorcamento',
            'codigoproposta',
            'proposta'
        ]
    },

    {
        apiName: 'locador',
        label: 'Locador',
        description: 'Nome do locador relacionado à renovação.',
        required: true,
        aliases: [
            'locador',
            'proprietario',
            'proprietarioimovel'
        ]
    },

    {
        apiName: 'locatario',
        label: 'Locatário',
        description: 'Nome do locatário ou segurado.',
        required: true,
        aliases: [
            'locatario',
            'inquilino',
            'segurado',
            'nomesegurado'
        ]
    },

    {
        apiName: 'produto',
        label: 'Produto',
        description: 'Produto relacionado à apólice.',
        required: false,
        aliases: [
            'produto',
            'tiposeguro',
            'tipodeseguro',
            'produtoseguradora'
        ]
    },

    {
        apiName: 'seguradora',
        label: 'Seguradora',
        description: 'Seguradora responsável pela apólice.',
        required: true,
        aliases: [
            'seguradora',
            'companhia',
            'companiaseguradora'
        ]
    },

    {
        apiName: 'imobiliaria',
        label: 'Imobiliária',
        description: 'Imobiliária relacionada à renovação.',
        required: false,
        aliases: [
            'imobiliaria',
            'nomeimobiliaria'
        ]
    },

    {
        apiName: 'dataFimVigencia',
        label: 'Data Fim Vigência',
        description: 'Data de término da vigência da apólice.',
        required: true,
        aliases: [
            'datafimvigencia',
            'fimvigencia',
            'finaldevigencia',
            'finalvigencia',
            'fimdoseguro',
            'datavencimento',
            'datadevencimentodaapolice',
            'datadevencdaapolice',
            'vencimentoapolice'
        ]
    }

];


export default class RenovacoesImportacaoExterna extends LightningElement {

    currentStep = '1';

    fileName = '';
    fileExtension = '';
    fileSizeLabel = '';

    errorMessage = '';

    isLoading = false;

    headers = [];

    allRows = [];
    previewRows = [];
    previewColumns = [];

    totalRowCount = 0;

    workbook = null;

    sheetNames = [];
    selectedSheetName = '';

    xlsxLibraryLoaded = false;
    xlsxLibraryPromise = null;

    mapping = {};
    suggestedFields = [];

    normalizedRows = [];
    confirmationPreviewRows = [];
    confirmationColumns = [];


    /* ============================================================ */
    /* ETAPAS                                                      */
    /* ============================================================ */

    get isUploadStep() {
        return this.currentStep === '1';
    }


    get isAnalysisStep() {
        return this.currentStep === '2';
    }


    get isPreviewStep() {
        return this.currentStep === '3';
    }


    get isMappingStep() {
        return this.currentStep === '4';
    }


    get isConfirmationStep() {
        return this.currentStep === '5';
    }


    /* ============================================================ */
    /* ARQUIVO                                                     */
    /* ============================================================ */

    get hasFile() {

        return Boolean(
            this.fileName &&
            this.headers.length > 0 &&
            this.totalRowCount > 0 &&
            !this.errorMessage
        );
    }


    get fileFormatLabel() {

        return this.fileExtension === 'xlsx'
            ? 'XLSX'
            : 'CSV';
    }


    get headerCount() {
        return this.headers.length;
    }


    get columnsBadgeLabel() {
        return `${this.headerCount} colunas`;
    }


    get previewLabel() {

        return `${this.previewRows.length} de ${this.totalRowCount} registros`;
    }


    /* ============================================================ */
    /* XLSX                                                        */
    /* ============================================================ */

    async ensureXlsxLibrary() {

        if (
            this.xlsxLibraryLoaded
        ) {
            return;
        }


        if (
            !this.xlsxLibraryPromise
        ) {

            this.xlsxLibraryPromise =
                loadScript(
                    this,
                    XLSX_RESOURCE
                )
                    .then(
                        () => {

                            if (
                                !window.XLSX
                            ) {

                                throw new Error(
                                    'A biblioteca XLSX não foi inicializada.'
                                );
                            }


                            this.xlsxLibraryLoaded =
                                true;
                        }
                    );
        }


        await this.xlsxLibraryPromise;
    }


    /* ============================================================ */
    /* UPLOAD                                                      */
    /* ============================================================ */

    async handleFileChange(event) {

        const files =
            event.target.files;


        if (
            !files ||
            files.length === 0
        ) {
            return;
        }


        const file =
            files[0];


        this.resetImportState();


        const extension =
            this.getFileExtension(
                file.name
            );


        if (
            extension !== 'csv' &&
            extension !== 'xlsx'
        ) {

            this.errorMessage =
                'Selecione um arquivo CSV ou XLSX.';

            return;
        }


        this.fileName =
            file.name;


        this.fileExtension =
            extension;


        this.fileSizeLabel =
            this.formatFileSize(
                file.size
            );


        this.isLoading =
            true;


        try {

            if (
                extension === 'csv'
            ) {

                await this.processCsvFile(
                    file
                );

            } else {

                await this.processXlsxFile(
                    file
                );
            }


        } catch (error) {

            console.error(
                'Erro ao processar arquivo:',
                error
            );


            this.errorMessage =
                error &&
                    error.message
                    ? error.message
                    : 'Não foi possível processar o arquivo selecionado.';


        } finally {

            this.isLoading =
                false;
        }
    }


    getFileExtension(fileName) {

        if (
            !fileName ||
            !fileName.includes('.')
        ) {
            return '';
        }


        return fileName
            .split('.')
            .pop()
            .toLowerCase();
    }


    formatFileSize(bytes) {

        if (
            !bytes
        ) {
            return '0 KB';
        }


        if (
            bytes < 1024
        ) {
            return `${bytes} bytes`;
        }


        const kb =
            bytes / 1024;


        if (
            kb < 1024
        ) {
            return `${kb.toFixed(1)} KB`;
        }


        const mb =
            kb / 1024;


        return `${mb.toFixed(1)} MB`;
    }


    /* ============================================================ */
    /* CSV                                                         */
    /* ============================================================ */

    async processCsvFile(file) {

        const content =
            await this.readTextFile(
                file,
                'UTF-8'
            );


        if (
            content &&
            content.includes('\uFFFD')
        ) {

            const windowsContent =
                await this.readTextFile(
                    file,
                    'windows-1252'
                );


            this.processCsvContent(
                windowsContent
            );


            return;
        }


        this.processCsvContent(
            content
        );
    }


    readTextFile(
        file,
        encoding
    ) {

        return new Promise(
            (
                resolve,
                reject
            ) => {

                const reader =
                    new FileReader();


                reader.onload =
                    () => {

                        resolve(
                            reader.result
                        );
                    };


                reader.onerror =
                    () => {

                        reject(
                            new Error(
                                'Não foi possível ler o arquivo.'
                            )
                        );
                    };


                reader.readAsText(
                    file,
                    encoding
                );
            }
        );
    }


    processCsvContent(content) {

        if (
            !content ||
            content.trim() === ''
        ) {

            throw new Error(
                'O arquivo selecionado está vazio.'
            );
        }


        const normalizedContent =
            content
                .replace(/^\uFEFF/, '')
                .replace(/\r\n/g, '\n')
                .replace(/\r/g, '\n');


        const firstLine =
            normalizedContent
                .split('\n')[0];


        const delimiter =
            this.detectDelimiter(
                firstLine
            );


        const matrix =
            this.parseCsv(
                normalizedContent,
                delimiter
            );


        this.loadMatrix(
            matrix
        );
    }


    detectDelimiter(firstLine) {

        const semicolonCount =
            this.countOutsideQuotes(
                firstLine,
                ';'
            );


        const commaCount =
            this.countOutsideQuotes(
                firstLine,
                ','
            );


        if (
            semicolonCount === 0 &&
            commaCount === 0
        ) {
            return ';';
        }


        return semicolonCount >= commaCount
            ? ';'
            : ',';
    }


    countOutsideQuotes(
        text,
        character
    ) {

        let count = 0;
        let insideQuotes = false;


        for (
            let index = 0;
            index < text.length;
            index++
        ) {

            const currentCharacter =
                text[index];


            if (
                currentCharacter === '"'
            ) {

                if (
                    insideQuotes &&
                    text[index + 1] === '"'
                ) {

                    index++;

                    continue;
                }


                insideQuotes =
                    !insideQuotes;


                continue;
            }


            if (
                currentCharacter === character &&
                !insideQuotes
            ) {

                count++;
            }
        }


        return count;
    }


    parseCsv(
        content,
        delimiter
    ) {

        const rows = [];

        let currentRow = [];
        let currentValue = '';
        let insideQuotes = false;


        for (
            let index = 0;
            index < content.length;
            index++
        ) {

            const character =
                content[index];


            if (
                character === '"'
            ) {

                if (
                    insideQuotes &&
                    content[index + 1] === '"'
                ) {

                    currentValue += '"';

                    index++;

                    continue;
                }


                insideQuotes =
                    !insideQuotes;


                continue;
            }


            if (
                character === delimiter &&
                !insideQuotes
            ) {

                currentRow.push(
                    currentValue
                );


                currentValue = '';


                continue;
            }


            if (
                character === '\n' &&
                !insideQuotes
            ) {

                currentRow.push(
                    currentValue
                );


                rows.push(
                    currentRow
                );


                currentRow = [];

                currentValue = '';


                continue;
            }


            currentValue +=
                character;
        }


        if (
            currentValue !== '' ||
            currentRow.length > 0
        ) {

            currentRow.push(
                currentValue
            );


            rows.push(
                currentRow
            );
        }


        return rows;
    }


    /* ============================================================ */
    /* XLSX                                                        */
    /* ============================================================ */

    async processXlsxFile(file) {

        await this.ensureXlsxLibrary();


        const buffer =
            await file.arrayBuffer();


        /*
         * IMPORTANTE:
         *
         * cellDates: true
         *
         * Faz com que células de data do Excel sejam carregadas
         * como Date em vez de texto formatado.
         */

        this.workbook =
            window.XLSX.read(
                buffer,
                {
                    type: 'array',
                    cellDates: true
                }
            );


        if (
            !this.workbook ||
            !this.workbook.SheetNames ||
            this.workbook.SheetNames.length === 0
        ) {

            throw new Error(
                'O arquivo XLSX não possui abas válidas.'
            );
        }


        this.sheetNames =
            [...this.workbook.SheetNames];


        this.selectedSheetName =
            this.findPreferredSheet(
                this.sheetNames
            );


        this.loadSelectedSheet();
    }


    findPreferredSheet(sheetNames) {

        if (
            !sheetNames ||
            sheetNames.length === 0
        ) {
            return '';
        }


        const preferred =
            sheetNames.find(
                sheetName => {

                    const normalized =
                        this.normalizeText(
                            sheetName
                        );


                    return (
                        normalized.includes(
                            'renovacao'
                        ) &&
                        !normalized.includes(
                            'original'
                        ) &&
                        !normalized.includes(
                            'basecompleta'
                        )
                    );
                }
            );


        if (
            preferred
        ) {
            return preferred;
        }


        const renewalSheet =
            sheetNames.find(
                sheetName =>
                    this.normalizeText(
                        sheetName
                    ).includes(
                        'renovacao'
                    )
            );


        return renewalSheet ||
            sheetNames[0];
    }


    loadSelectedSheet() {

        if (
            !this.workbook ||
            !this.selectedSheetName
        ) {
            return;
        }


        const worksheet =
            this.workbook.Sheets[
            this.selectedSheetName
            ];


        if (
            !worksheet
        ) {

            throw new Error(
                'Não foi possível abrir a aba selecionada.'
            );
        }


        /*
         * IMPORTANTE:
         *
         * raw: true
         *
         * Não utiliza o texto visual formatado da célula.
         * Preserva Date, número e texto conforme o Excel.
         */

        const matrix =
            window.XLSX.utils.sheet_to_json(
                worksheet,
                {
                    header: 1,
                    raw: true,
                    defval: '',
                    blankrows: false
                }
            );


        this.loadMatrix(
            matrix
        );
    }


    /* ============================================================ */
    /* MATRIZ                                                      */
    /* ============================================================ */

    loadMatrix(matrix) {

        const nonEmptyRows =
            (matrix || [])
                .map(
                    row =>
                        Array.isArray(row)
                            ? row
                            : []
                )
                .filter(
                    row =>
                        row.some(
                            value =>
                                !this.isBlank(
                                    value
                                )
                        )
                );


        if (
            nonEmptyRows.length === 0
        ) {

            throw new Error(
                'O arquivo não possui dados.'
            );
        }


        if (
            nonEmptyRows.length === 1
        ) {

            throw new Error(
                'O arquivo possui cabeçalho, mas não possui registros.'
            );
        }


        const headerValues =
            nonEmptyRows[0]
                .map(
                    value =>
                        String(
                            value === null ||
                                value === undefined
                                ? ''
                                : value
                        ).trim()
                );


        if (
            headerValues.some(
                label =>
                    !label
            )
        ) {

            throw new Error(
                'Existe uma ou mais colunas sem nome no cabeçalho.'
            );
        }


        this.headers =
            headerValues.map(
                (
                    label,
                    index
                ) => ({

                    key:
                        `col${index}`,

                    label,

                    normalized:
                        this.normalizeText(
                            label
                        )
                })
            );


        this.previewColumns =
            this.headers.map(
                header => ({

                    label:
                        header.label,

                    fieldName:
                        header.key,

                    type:
                        'text',

                    wrapText:
                        false
                })
            );


        const dataMatrix =
            nonEmptyRows.slice(1);


        this.allRows =
            dataMatrix.map(
                (
                    values,
                    rowIndex
                ) => {

                    const row = {

                        _rowId:
                            String(
                                rowIndex + 1
                            )
                    };


                    this.headers.forEach(
                        (
                            header,
                            columnIndex
                        ) => {

                            const value =
                                values[
                                columnIndex
                                ];


                            row[
                                header.key
                            ] =
                                this.formatCellValue(
                                    value
                                );
                        }
                    );


                    return row;
                }
            );


        this.totalRowCount =
            this.allRows.length;


        this.previewRows =
            this.allRows.slice(
                0,
                PREVIEW_LIMIT
            );


        this.errorMessage =
            '';


        this.autoMapColumns();
    }


    /* ============================================================ */
    /* FORMATAÇÃO DE CÉLULAS                                       */
    /* ============================================================ */

    formatCellValue(value) {

        if (
            value === null ||
            value === undefined
        ) {

            return '';
        }


        /*
         * DATA NATIVA DO XLSX
         */

        if (
            value instanceof Date &&
            !Number.isNaN(
                value.getTime()
            )
        ) {

            return this.formatDateForDisplay(
                value
            );
        }


        return String(
            value
        ).trim();
    }


    formatDateForDisplay(date) {

        const day =
            String(
                date.getDate()
            ).padStart(
                2,
                '0'
            );


        const month =
            String(
                date.getMonth() + 1
            ).padStart(
                2,
                '0'
            );


        const year =
            date.getFullYear();


        return `${day}/${month}/${year}`;
    }


    /* ============================================================ */
    /* NORMALIZAÇÃO / MAPEAMENTO                                   */
    /* ============================================================ */

    normalizeText(text) {

        return String(
            text || ''
        )
            .normalize('NFD')
            .replace(
                /[\u0300-\u036f]/g,
                ''
            )
            .toLowerCase()
            .replace(
                /[^a-z0-9]/g,
                ''
            );
    }


    findHeaderByAliases(aliases) {

        const normalizedAliases =
            aliases.map(
                alias =>
                    this.normalizeText(
                        alias
                    )
            );


        let found =
            this.headers.find(
                header =>
                    normalizedAliases.includes(
                        header.normalized
                    )
            );


        if (
            found
        ) {
            return found;
        }


        found =
            this.headers.find(
                header =>
                    normalizedAliases.some(
                        alias =>
                            header.normalized.includes(
                                alias
                            ) ||
                            alias.includes(
                                header.normalized
                            )
                    )
            );


        return found ||
            null;
    }


    resolveSourceColumn(processFieldName) {

        const mapped =
            this.mapping[
            processFieldName
            ];


        if (
            mapped &&
            this.headers.some(
                header =>
                    header.key ===
                    mapped
            )
        ) {

            return mapped;
        }


        const field =
            PROCESS_FIELDS.find(
                item =>
                    item.apiName ===
                    processFieldName
            );


        if (
            !field
        ) {
            return '';
        }


        const header =
            this.findHeaderByAliases(
                field.aliases
            );


        return header
            ? header.key
            : '';
    }


    autoMapColumns() {

        const mapping =
            {};


        const used =
            new Set();


        PROCESS_FIELDS.forEach(
            field => {

                const header =
                    this.findHeaderByAliases(
                        field.aliases
                    );


                if (
                    header &&
                    !used.has(
                        header.key
                    )
                ) {

                    mapping[
                        field.apiName
                    ] =
                        header.key;


                    used.add(
                        header.key
                    );
                }
            }
        );


        this.mapping =
            mapping;


        this.suggestedFields =
            Object.keys(
                mapping
            );
    }


    /* ============================================================ */
    /* CAMPOS OBRIGATÓRIOS                                         */
    /* ============================================================ */

    get requiredFields() {

        return PROCESS_FIELDS.filter(
            field =>
                field.required
        );
    }


    get requiredTotal() {

        return this.requiredFields.length;
    }


    get requiredMappedCount() {

        return this.requiredFields.filter(
            field =>
                Boolean(
                    this.resolveSourceColumn(
                        field.apiName
                    )
                )
        ).length;
    }


    get requiredDetectionLabel() {

        return `${this.requiredMappedCount}/${this.requiredTotal}`;
    }


    get hasMissingRequiredMappings() {

        return (
            this.requiredMappedCount <
            this.requiredTotal
        );
    }


    get missingRequiredFields() {

        return this.requiredFields.filter(
            field =>
                !this.resolveSourceColumn(
                    field.apiName
                )
        );
    }


    get missingRequiredLabels() {

        return this.missingRequiredFields
            .map(
                field =>
                    field.label
            )
            .join(', ');
    }


    get requiredCardClass() {

        return this.hasMissingRequiredMappings
            ? 'kpi-card kpi-card-warning'
            : 'kpi-card kpi-card-success';
    }


    /* ============================================================ */
    /* DISTRIBUIÇÕES                                                */
    /* ============================================================ */

    get seguradoraDistribution() {

        return this.buildValueDistribution(
            'seguradora'
        );
    }


    get imobiliariaDistribution() {

        return this.buildValueDistribution(
            'imobiliaria'
        );
    }


    get monthDistribution() {

        const sourceColumn =
            this.resolveSourceColumn(
                'dataFimVigencia'
            );


        if (
            !sourceColumn
        ) {
            return [];
        }


        const counters =
            new Map();


        let invalidCount =
            0;


        this.allRows.forEach(
            row => {

                const value =
                    row[
                    sourceColumn
                    ];


                if (
                    this.isBlank(
                        value
                    )
                ) {

                    invalidCount++;

                    return;
                }


                const monthInfo =
                    this.getMonthInfo(
                        value
                    );


                if (
                    !monthInfo
                ) {

                    invalidCount++;

                    return;
                }


                if (
                    !counters.has(
                        monthInfo.key
                    )
                ) {

                    counters.set(
                        monthInfo.key,
                        {
                            key:
                                monthInfo.key,

                            label:
                                monthInfo.label,

                            count:
                                0,

                            sortValue:
                                monthInfo.sortValue
                        }
                    );
                }


                counters.get(
                    monthInfo.key
                ).count++;
            }
        );


        const result =
            Array.from(
                counters.values()
            )
                .sort(
                    (
                        a,
                        b
                    ) =>
                        a.sortValue -
                        b.sortValue
                );


        if (
            invalidCount > 0
        ) {

            result.push(
                {
                    key:
                        '__INVALID_DATE__',

                    label:
                        'Sem data válida',

                    count:
                        invalidCount,

                    sortValue:
                        Number.MAX_SAFE_INTEGER
                }
            );
        }


        return result;
    }


    buildValueDistribution(
        processField
    ) {

        const sourceColumn =
            this.resolveSourceColumn(
                processField
            );


        if (
            !sourceColumn
        ) {
            return [];
        }


        const counters =
            new Map();


        let blankCount =
            0;


        this.allRows.forEach(
            row => {

                const rawValue =
                    row[
                    sourceColumn
                    ];


                if (
                    this.isBlank(
                        rawValue
                    )
                ) {

                    blankCount++;

                    return;
                }


                const displayValue =
                    String(
                        rawValue
                    )
                        .trim()
                        .replace(
                            /\s+/g,
                            ' '
                        );


                const normalizedValue =
                    this.normalizeText(
                        displayValue
                    );


                if (
                    !counters.has(
                        normalizedValue
                    )
                ) {

                    counters.set(
                        normalizedValue,
                        {
                            key:
                                normalizedValue,

                            label:
                                displayValue,

                            count:
                                0
                        }
                    );
                }


                counters.get(
                    normalizedValue
                ).count++;
            }
        );


        if (
            counters.size === 0
        ) {
            return [];
        }


        if (
            blankCount > 0
        ) {

            counters.set(
                '__BLANK__',
                {
                    key:
                        '__BLANK__',

                    label:
                        'Não informada',

                    count:
                        blankCount
                }
            );
        }


        return Array.from(
            counters.values()
        )
            .sort(
                (
                    a,
                    b
                ) => {

                    if (
                        b.count !==
                        a.count
                    ) {

                        return (
                            b.count -
                            a.count
                        );
                    }


                    return a.label.localeCompare(
                        b.label,
                        'pt-BR'
                    );
                }
            );
    }


    get hasSeguradoraDistribution() {

        return this.seguradoraDistribution.length >
            0;
    }


    get hasImobiliariaDistribution() {

        return this.imobiliariaDistribution.length >
            0;
    }


    get hasMonthDistribution() {

        return this.monthDistribution.length >
            0;
    }


    get seguradoraSummaryLabel() {

        const count =
            this.seguradoraDistribution.filter(
                item =>
                    item.key !==
                    '__BLANK__'
            ).length;


        return count === 1
            ? '1 seguradora'
            : `${count} seguradoras`;
    }


    get imobiliariaSummaryLabel() {

        const count =
            this.imobiliariaDistribution.filter(
                item =>
                    item.key !==
                    '__BLANK__'
            ).length;


        return count === 1
            ? '1 imobiliária'
            : `${count} imobiliárias`;
    }


    get monthSummaryLabel() {

        const count =
            this.monthDistribution.filter(
                item =>
                    item.key !==
                    '__INVALID_DATE__'
            ).length;


        return count === 1
            ? '1 mês'
            : `${count} meses`;
    }


    get monthDistributionMessage() {

        if (
            !this.resolveSourceColumn(
                'dataFimVigencia'
            )
        ) {

            return 'Data Fim Vigência não foi identificada no arquivo.';
        }


        return 'Não foi possível identificar meses válidos.';
    }


    /* ============================================================ */
    /* LEITURA DE DATAS                                            */
    /* ============================================================ */

    buildMonthInfo(
        year,
        month
    ) {

        const key =
            `${year}-${String(
                month
            ).padStart(
                2,
                '0'
            )}`;


        return {

            key,

            label:
                `${MONTH_NAMES[month - 1]}/${year}`,

            sortValue:
                (
                    year *
                    100
                ) +
                month
        };
    }


    getMonthInfo(value) {

        if (
            value === null ||
            value === undefined
        ) {
            return null;
        }


        /*
         * Caso algum Date chegue diretamente até aqui.
         */

        if (
            value instanceof Date &&
            !Number.isNaN(
                value.getTime()
            )
        ) {

            return this.buildMonthInfo(
                value.getFullYear(),
                value.getMonth() + 1
            );
        }


        const text =
            String(
                value
            ).trim();


        if (
            !text
        ) {
            return null;
        }


        let match;
        let year;
        let month;
        let day;


        /* ======================================================== */
        /* DD/MM/AAAA                                               */
        /* ======================================================== */

        match =
            text.match(
                /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
            );


        if (
            match
        ) {

            day =
                Number(
                    match[1]
                );


            month =
                Number(
                    match[2]
                );


            year =
                Number(
                    match[3]
                );


            if (
                this.isValidCalendarDate(
                    year,
                    month,
                    day
                )
            ) {

                return this.buildMonthInfo(
                    year,
                    month
                );
            }
        }


        /* ======================================================== */
        /* AAAA-MM-DD                                               */
        /* ======================================================== */

        match =
            text.match(
                /^(\d{4})-(\d{1,2})-(\d{1,2})/
            );


        if (
            match
        ) {

            year =
                Number(
                    match[1]
                );


            month =
                Number(
                    match[2]
                );


            day =
                Number(
                    match[3]
                );


            if (
                this.isValidCalendarDate(
                    year,
                    month,
                    day
                )
            ) {

                return this.buildMonthInfo(
                    year,
                    month
                );
            }
        }


        /* ======================================================== */
        /* SERIAL EXCEL                                             */
        /* ======================================================== */

        if (
            /^\d+(?:[.,]\d+)?$/.test(
                text
            )
        ) {

            const serial =
                Number(
                    text.replace(
                        ',',
                        '.'
                    )
                );


            if (
                serial > 20000 &&
                serial < 100000
            ) {

                const excelEpoch =
                    new Date(
                        Date.UTC(
                            1899,
                            11,
                            30
                        )
                    );


                const milliseconds =
                    Math.floor(
                        serial
                    ) *
                    86400000;


                const date =
                    new Date(
                        excelEpoch.getTime() +
                        milliseconds
                    );


                return this.buildMonthInfo(
                    date.getUTCFullYear(),
                    date.getUTCMonth() + 1
                );
            }
        }


        return null;
    }


    isValidCalendarDate(
        year,
        month,
        day
    ) {

        if (
            !year ||
            !month ||
            !day ||
            month < 1 ||
            month > 12 ||
            day < 1 ||
            day > 31
        ) {

            return false;
        }


        const date =
            new Date(
                year,
                month - 1,
                day
            );


        return (
            date.getFullYear() ===
            year &&
            date.getMonth() ===
            month - 1 &&
            date.getDate() ===
            day
        );
    }


    /* ============================================================ */
    /* NAVEGAÇÃO                                                   */
    /* ============================================================ */

    handleGoToAnalysis() {

        if (
            !this.hasFile
        ) {
            return;
        }


        this.currentStep =
            '2';
    }


    handleBackToUpload() {

        this.currentStep =
            '1';
    }


    handleGoToPreview() {

        this.currentStep =
            '3';
    }


    handleBackToAnalysis() {

        this.currentStep =
            '2';
    }


    handleGoToMapping() {

        this.currentStep =
            '4';
    }


    handleBackToPreview() {

        this.currentStep =
            '3';
    }


    handleBackToMapping() {

        this.currentStep =
            '4';
    }


    /* ============================================================ */
    /* DE/PARA                                                     */
    /* ============================================================ */

    get headerOptions() {

        return this.headers.map(
            header => ({
                label:
                    header.label,

                value:
                    header.key
            })
        );
    }


    get optionalHeaderOptions() {

        return [
            {
                label:
                    'Não mapear',

                value:
                    NONE_VALUE
            },

            ...this.headerOptions
        ];
    }


    get mappingRows() {

        const usage =
            {};


        Object.values(
            this.mapping
        ).forEach(
            key => {

                if (
                    key
                ) {

                    usage[
                        key
                    ] =
                        (
                            usage[
                            key
                            ] ||
                            0
                        ) +
                        1;
                }
            }
        );


        const firstRow =
            this.allRows[0] ||
            {};


        return PROCESS_FIELDS.map(
            field => {

                const sourceColumn =
                    this.mapping[
                    field.apiName
                    ] ||
                    '';


                const hasValue =
                    Boolean(
                        sourceColumn
                    );


                const isDuplicate =
                    hasValue &&
                    usage[
                    sourceColumn
                    ] >
                    1;


                const sample =
                    hasValue
                        ? firstRow[
                        sourceColumn
                        ]
                        : '';


                let hint =
                    '';


                if (
                    isDuplicate
                ) {

                    hint =
                        'Esta coluna está sendo usada em mais de um campo.';

                } else if (
                    hasValue
                ) {

                    hint =
                        sample
                            ? `Exemplo: ${sample}`
                            : 'A primeira linha está vazia nesta coluna.';
                }


                return {

                    apiName:
                        field.apiName,

                    label:
                        field.label,

                    description:
                        field.description,

                    required:
                        field.required,

                    value:
                        field.required
                            ? sourceColumn
                            : (
                                sourceColumn ||
                                NONE_VALUE
                            ),

                    options:
                        field.required
                            ? this.headerOptions
                            : this.optionalHeaderOptions,

                    placeholder:
                        field.required
                            ? 'Selecione uma coluna'
                            : 'Não mapear',

                    isMapped:
                        hasValue &&
                        !isDuplicate,

                    isDuplicate,

                    isPending:
                        !hasValue,

                    isSuggested:
                        hasValue &&
                        !isDuplicate &&
                        this.suggestedFields.includes(
                            field.apiName
                        ),

                    hint,

                    hintClass:
                        isDuplicate
                            ? 'mapping-hint mapping-hint-error'
                            : 'mapping-hint',

                    rowClass:
                        isDuplicate
                            ? 'mapping-row mapping-row-error'
                            : 'mapping-row'
                };
            }
        );
    }


    get hasDuplicateMapping() {

        return this.mappingRows.some(
            row =>
                row.isDuplicate
        );
    }


    get mappedRequiredCount() {

        return this.mappingRows.filter(
            row =>
                row.required &&
                row.isMapped
        ).length;
    }


    get mappingProgressLabel() {

        return `${this.mappedRequiredCount} de ${this.requiredTotal} obrigatórios`;
    }


    get isConfirmationDisabled() {

        return (
            this.mappedRequiredCount <
            this.requiredTotal
        ) ||
            this.hasDuplicateMapping;
    }


    get mappingStatusLabel() {

        if (
            this.hasDuplicateMapping
        ) {

            return 'Existe uma coluna utilizada mais de uma vez.';
        }


        const missing =
            this.requiredTotal -
            this.mappedRequiredCount;


        if (
            missing === 0
        ) {

            return 'Tudo pronto para revisar a confirmação.';
        }


        return missing === 1
            ? 'Falta 1 campo obrigatório.'
            : `Faltam ${missing} campos obrigatórios.`;
    }


    get mappingStatusClass() {

        if (
            this.hasDuplicateMapping
        ) {

            return 'mapping-status-text mapping-status-error';
        }


        return this.isConfirmationDisabled
            ? 'mapping-status-text'
            : 'mapping-status-text mapping-status-ready';
    }


    handleMappingChange(event) {

        const field =
            event.target.dataset.field;


        let value =
            event.detail.value;


        if (
            value === NONE_VALUE
        ) {

            value =
                '';
        }


        this.mapping = {

            ...this.mapping,

            [field]:
                value
        };


        this.suggestedFields =
            this.suggestedFields.filter(
                fieldName =>
                    fieldName !==
                    field
            );
    }


    /* ============================================================ */
    /* CONFIRMAÇÃO                                                 */
    /* ============================================================ */

    handleGoToConfirmation() {

        if (
            this.isConfirmationDisabled
        ) {
            return;
        }


        this.buildConfirmationData();


        this.currentStep =
            '5';
    }


    buildConfirmationData() {

        this.normalizedRows =
            this.allRows.map(
                row => {

                    const normalizedRow = {

                        _rowId:
                            row._rowId,

                        numeroOrcamento:
                            this.getMappedValue(
                                row,
                                'numeroOrcamento'
                            ),

                        locador:
                            this.getMappedValue(
                                row,
                                'locador'
                            ),

                        locatario:
                            this.getMappedValue(
                                row,
                                'locatario'
                            ),

                        produto:
                            this.getMappedValue(
                                row,
                                'produto'
                            ),

                        seguradora:
                            this.getMappedValue(
                                row,
                                'seguradora'
                            ),

                        imobiliaria:
                            this.getMappedValue(
                                row,
                                'imobiliaria'
                            ),

                        dataFimVigencia:
                            this.getMappedValue(
                                row,
                                'dataFimVigencia'
                            )
                    };


                    const errors =
                        this.validateNormalizedRow(
                            normalizedRow
                        );


                    return {

                        ...normalizedRow,

                        validationStatus:
                            errors.length ===
                                0
                                ? 'Válido'
                                : 'Inválido',

                        validationMessage:
                            errors.length ===
                                0
                                ? '-'
                                : errors.join(
                                    '; '
                                )
                    };
                }
            );


        this.confirmationPreviewRows =
            this.normalizedRows.slice(
                0,
                PREVIEW_LIMIT
            );


        this.confirmationColumns = [

            {
                label:
                    'Nº orçamento',

                fieldName:
                    'numeroOrcamento',

                type:
                    'text'
            },

            {
                label:
                    'Locador',

                fieldName:
                    'locador',

                type:
                    'text'
            },

            {
                label:
                    'Locatário',

                fieldName:
                    'locatario',

                type:
                    'text'
            },

            {
                label:
                    'Produto',

                fieldName:
                    'produto',

                type:
                    'text'
            },

            {
                label:
                    'Seguradora',

                fieldName:
                    'seguradora',

                type:
                    'text'
            },

            {
                label:
                    'Imobiliária',

                fieldName:
                    'imobiliaria',

                type:
                    'text'
            },

            {
                label:
                    'Data Fim Vigência',

                fieldName:
                    'dataFimVigencia',

                type:
                    'text'
            },

            {
                label:
                    'Status',

                fieldName:
                    'validationStatus',

                type:
                    'text'
            },

            {
                label:
                    'Erros',

                fieldName:
                    'validationMessage',

                type:
                    'text',

                wrapText:
                    true
            }

        ];
    }


    getMappedValue(
        row,
        processField
    ) {

        const sourceColumn =
            this.mapping[
            processField
            ];


        if (
            !sourceColumn
        ) {

            return '';
        }


        return row[
            sourceColumn
        ] ||
            '';
    }


    validateNormalizedRow(row) {

        const errors =
            [];


        if (
            this.isBlank(
                row.numeroOrcamento
            )
        ) {

            errors.push(
                'Nº orçamento não preenchido'
            );
        }


        if (
            this.isBlank(
                row.locador
            )
        ) {

            errors.push(
                'Locador não preenchido'
            );
        }


        if (
            this.isBlank(
                row.locatario
            )
        ) {

            errors.push(
                'Locatário não preenchido'
            );
        }


        if (
            this.isBlank(
                row.seguradora
            )
        ) {

            errors.push(
                'Seguradora não preenchida'
            );
        }


        if (
            this.isBlank(
                row.dataFimVigencia
            )
        ) {

            errors.push(
                'Data Fim Vigência não preenchida'
            );
        }


        return errors;
    }


    get validRowCount() {

        return this.normalizedRows.filter(
            row =>
                row.validationStatus ===
                'Válido'
        ).length;
    }


    get invalidRowCount() {

        return this.normalizedRows.filter(
            row =>
                row.validationStatus ===
                'Inválido'
        ).length;
    }


    get hasInvalidRows() {

        return this.invalidRowCount >
            0;
    }


    /* ============================================================ */
    /* UTILITÁRIOS                                                 */
    /* ============================================================ */

    isBlank(value) {

        return (
            value === null ||
            value === undefined ||
            String(
                value
            ).trim() === ''
        );
    }


    resetImportState() {

        this.currentStep =
            '1';


        this.fileName =
            '';


        this.fileExtension =
            '';


        this.fileSizeLabel =
            '';


        this.errorMessage =
            '';


        this.headers =
            [];


        this.allRows =
            [];


        this.previewRows =
            [];


        this.previewColumns =
            [];


        this.totalRowCount =
            0;


        this.workbook =
            null;


        this.sheetNames =
            [];


        this.selectedSheetName =
            '';


        this.mapping =
            {};


        this.suggestedFields =
            [];


        this.normalizedRows =
            [];


        this.confirmationPreviewRows =
            [];


        this.confirmationColumns =
            [];
    }
}
