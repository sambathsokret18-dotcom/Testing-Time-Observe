/* ==========================================
    FULL SCRIPT.JS - TIME STUDY, DASHBOARD & GRAPH SYSTEM (FINAL INTEGRATED)
    - Added Checkmark Badge & Average CT before SMV
    - Styled UI Buttons for Process View
========================================== */

document.addEventListener('DOMContentLoaded', () => {
    
    // --- ១. មុខងារប្ដូរទំព័រពេលចុចលើ Sidebar Menu ---
    const menuLinks = document.querySelectorAll('.sidebar ul li a');
    const pages = document.querySelectorAll('.page-content');

    menuLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            
            menuLinks.forEach(item => item.classList.remove('active'));
            this.classList.add('active');

            const targetPageId = this.getAttribute('data-page');

            pages.forEach(page => {
                if (page.id === targetPageId) {
                    page.style.display = 'block';
                    if (targetPageId === 'time-page') {
                        ['section-page', 'style-page', 'rampup-page', 'prep-page', 'process-page', 'timestudy-page'].forEach(pId => {
                            const p = document.getElementById(pId);
                            if (p) p.style.display = 'none';
                        });
                        renderMasterSectList();
                    } else if (targetPageId === 'dashboard-page') {
                        renderDashboardStats();
                    }
                } else {
                    page.style.display = 'none';
                }
            });
        });
    });

    // --- ២. មុខងារគ្រប់គ្រងការ Upload និង Excel Processing ---
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('excel-file');
    const fileInfo = document.getElementById('file-info');
    const fileNameSpan = document.getElementById('file-name');
    const removeFileBtn = document.getElementById('remove-file');
    const uploadForm = document.getElementById('upload-form');

    let uploadedFilesList = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];

    renderTable();
    updateTotalFilesCard();

    if (dropZone && fileInput) {
        dropZone.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                validateAndProcessFile(e.target.files[0]);
            }
        });

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '#2563eb';
            dropZone.style.backgroundColor = '#eff6ff';
        });

        dropZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '#cbd5e1';
            dropZone.style.backgroundColor = '#f8fafc';
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            if (e.dataTransfer.files.length > 0) {
                validateAndProcessFile(e.dataTransfer.files[0]);
            }
        });

        function validateAndProcessFile(file) {
            const fileName = file.name.toLowerCase();
            
            if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
                showCustomAlert('សូមអភ័យទោស!', 'ប្រព័ន្ធអនុញ្ញាតឱ្យ Upload តែឯកសារ Excel (.xlsx ឬ .xls) ប៉ុណ្ណោះ។', 'fa-solid fa-triangle-exclamation', '#f59e0b');
                resetFileInput();
                return;
            }

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    if (jsonData.length === 0) {
                        showCustomAlert('ទម្រង់មិនត្រឹមត្រូវ', 'ហ្វាល Excel នេះទទេ គ្មានទិន្នន័យទេ!', 'fa-solid fa-triangle-exclamation', '#dc2626');
                        resetFileInput();
                        return;
                    }

                    const fileHeaders = jsonData[0].map(h => String(h).trim().toLowerCase());
                    const requiredHeaders = ['title', 'primarykey', 'smv'];
                    
                    const missingHeaders = requiredHeaders.filter(req => 
                        !fileHeaders.some(fh => fh.includes(req))
                    );

                    if (missingHeaders.length > 0) {
                        showCustomAlert('ខុសទម្រង់ក្បាលតារាង', 'ហ្វាល Excel ខ្វះក្បាលតារាងចាំបាច់ (Title, PrimaryKey, SMV)', 'fa-solid fa-triangle-exclamation', '#dc2626');
                        resetFileInput();
                    } else {
                        fileInfo.style.display = 'flex';
                        dropZone.style.display = 'none';
                        fileNameSpan.textContent = file.name;
                    }
                } catch (error) {
                    console.error("Excel Error:", error);
                    showCustomAlert('កំហុសឆ្គង', 'មិនអាចអានហ្វាល Excel នេះបានទេ: ' + error.message, 'fa-solid fa-triangle-exclamation', '#dc2626');
                    resetFileInput();
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function resetFileInput() {
            fileInput.value = '';
            fileInfo.style.display = 'none';
            dropZone.style.display = 'flex';
        }

        if (removeFileBtn) {
            removeFileBtn.addEventListener('click', () => resetFileInput());
        }

        if (uploadForm) {
            uploadForm.addEventListener('submit', (e) => {
                e.preventDefault();
                if (fileInput.files.length === 0) {
                    showCustomAlert('ជូនដំណឹង', 'សូមជ្រើសរើសឯកសារ Excel ជាមុនសិន!', 'fa-solid fa-circle-info', '#3b82f6');
                    return;
                }

                const file = fileInput.files[0];
                const reader = new FileReader();

                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                        
                        const rawArrayData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        const rawHeaders = rawArrayData.length > 0 ? rawArrayData[0].map(h => String(h).trim()) : [];

                        const jsonData = XLSX.utils.sheet_to_json(worksheet);

                        const titlesSet = new Set();
                        const sectionsMap = {};
                        const stylesMap = {};
                        const rampupMap = {};
                        const prepMap = {};

                        jsonData.forEach((row, index) => {
                            const getRowVal = (keys) => {
                                const foundKey = Object.keys(row).find(k => keys.some(kw => k.trim().toLowerCase() === kw.toLowerCase()));
                                return foundKey && row[foundKey] !== undefined ? String(row[foundKey]).trim() : '';
                            };

                            const titleVal = getRowVal(['title', 'factory']) || 'General Factory';
                            const sectionVal = getRowVal(['section']) || 'General Section';
                            const styleVal = getRowVal(['style', 'stylename']) || 'General Style';
                            const rampupVal = getRowVal(['rampupd', 'rampupday', 'ramp_up_day', 'rampup', 'day']) || 'Day 1';
                            const primaryKeyVal = getRowVal(['primarykey', 'primary_key']) || 'N/A';
                            const sewerNameVal = getRowVal(['sewername', 'sewer_name', 'sewer', 'sewer name']) || 'No Sewer';
                            const prepAssemblyVal = getRowVal(['prep_assembly', 'prepassembly', 'prep', 'assembly', 'prep/assembly']) || 'Detail Item';
                            
                            const processKhmerVal = getRowVal(['process(region)', 'processregion', 'process_region', 'process_khmer', 'khmer']) || '';
                            const processThaiVal = getRowVal(['process', 'processname', 'operation', 'operationname', 'thai']) || '';

                            let finalProcessRegion = processKhmerVal ? processKhmerVal : (processThaiVal || 'General Process');
                            let finalProcessName = processThaiVal ? processThaiVal : processKhmerVal;
                            
                            const noVal = getRowVal(['no', 'number']) || String(index + 1);
                            const seqVal = getRowVal(['seq', 'sequence', 'opseq']) || '99';
                            const mcVal = getRowVal(['mc', 'machine', 'machinename', 'm/c']) || '-';
                            const weightVal = getRowVal(['weight', 'wt']) || '-';
                            const dateTimeStartVal = getRowVal(['datetime-start', 'datetimestart', 'start_time']) || '';

                            let rawSmv = parseFloat(getRowVal(['smv', 'ngiesmv'])) || 0;
                            let formattedSmv = !isNaN(rawSmv) ? rawSmv.toFixed(2) : '0.00';

                            titlesSet.add(titleVal);

                            if (!sectionsMap[titleVal]) sectionsMap[titleVal] = new Set();
                            sectionsMap[titleVal].add(sectionVal);

                            const styleCompositeKey = `${titleVal}_${sectionVal}`;
                            if (!stylesMap[styleCompositeKey]) stylesMap[styleCompositeKey] = new Set();
                            stylesMap[styleCompositeKey].add(styleVal);

                            const rampupCompositeKey = `${titleVal}_${sectionVal}_${styleVal}`;
                            if (!rampupMap[rampupCompositeKey]) rampupMap[rampupCompositeKey] = new Set();
                            rampupMap[rampupCompositeKey].add(rampupVal);

                            const prepCompositeKey = `${titleVal}_${sectionVal}_${styleVal}_${rampupVal}`;
                            if (!prepMap[prepCompositeKey]) prepMap[prepCompositeKey] = [];

                            prepMap[prepCompositeKey].push({
                                id: 'EXCEL_ROW_' + Date.now() + '_' + index,
                                title: titleVal,
                                factory: titleVal,
                                section: sectionVal,
                                styleName: styleVal,
                                rampupDay: rampupVal,
                                primaryKey: primaryKeyVal,
                                sewerName: sewerNameVal,
                                prepAssembly: prepAssemblyVal,
                                processName: finalProcessName,
                                processRegion: finalProcessRegion,
                                smv: formattedSmv,
                                no: noVal,
                                seq: seqVal,
                                operatorSeq: parseInt(seqVal) || 99,
                                mc: mcVal,
                                weight: weightVal,
                                dateTimeStart: dateTimeStartVal,
                                rawRowData: row
                            });
                        });

                        const uniqueTitles = Array.from(titlesSet);
                        const formattedSectionsMap = {};
                        for (let t in sectionsMap) formattedSectionsMap[t] = Array.from(sectionsMap[t]);

                        const formattedStylesMap = {};
                        for (let k in stylesMap) formattedStylesMap[k] = Array.from(stylesMap[k]);

                        const formattedRampupMap = {};
                        for (let k in rampupMap) formattedRampupMap[k] = Array.from(rampupMap[k]);

                        const now = new Date();
                        const dateString = now.toLocaleDateString('km-KH') + ' ' + now.toLocaleTimeString('km-KH', {hour: '2-digit', minute:'2-digit'});

                        uploadedFilesList.unshift({
                            user: 'SOKRET',
                            name: file.name,
                            date: dateString,
                            excelHeaders: rawHeaders,
                            titles: uniqueTitles,
                            sectionsMap: formattedSectionsMap,
                            stylesMap: formattedStylesMap,
                            rampupMap: formattedRampupMap,
                            prepMap: prepMap
                        });

                        localStorage.setItem('uploadedExcelFiles', JSON.stringify(uploadedFilesList));
                        renderTable();
                        updateTotalFilesCard();
                        showCustomAlert('ជោគជ័យ', 'ឯកសារ Excel ["' + file.name + '"] ត្រូវបាន Upload ជោគជ័យ!', 'fa-solid fa-circle-check', '#10b981');
                        resetFileInput();
                    } catch (err) {
                        console.error("Upload Error:", err);
                        showCustomAlert('កំហុសឆ្គង', 'មិនអាចអានទិន្នន័យពីហ្វាល Excel នេះបានទេ: ' + err.message, 'fa-solid fa-triangle-exclamation', '#dc2626');
                    }
                };
                reader.readAsArrayBuffer(file);
            });
        }
    }

    function renderTable() {
        const tableBody = document.getElementById('files-table-body');
        const fileCountSpan = document.getElementById('file-count');
        if (fileCountSpan) fileCountSpan.textContent = uploadedFilesList.length;
        if (!tableBody) return;

        if (uploadedFilesList.length === 0) {
            tableBody.innerHTML = `<tr id="empty-row"><td colspan="7" style="text-align: center; color: #9ca3af;">មិនទាន់មានឯកសារត្រូវបាន Upload នៅឡើយទេ</td></tr>`;
            return;
        }

        let html = '';
        uploadedFilesList.forEach((item, index) => {
            html += `
                <tr>
                    <td>${index + 1}</td>
                    <td>${item.user || 'SOKRET'}</td>
                    <td><i class="fa-solid fa-file-excel" style="color: #10b981; margin-right: 8px;"></i> ${item.name}</td>
                    <td>${item.date}</td>
                    <td><span class="status-badge">Done</span></td>
                    <td><button class="action-btn download-btn-action" data-filename="${item.name}"><i class="fa-solid fa-download"></i> Download</button></td>
                    <td><button class="action-btn delete-btn-action" data-index="${index}"><i class="fa-solid fa-trash"></i> Delete</button></td>
                </tr>
            `;
        });
        tableBody.innerHTML = html;

        tableBody.querySelectorAll('.delete-btn-action').forEach(btn => {
            btn.addEventListener('click', function() {
                deleteFile(parseInt(this.getAttribute('data-index')));
            });
        });

        tableBody.querySelectorAll('.download-btn-action').forEach(btn => {
            btn.addEventListener('click', function() {
                downloadFile(this.getAttribute('data-filename'));
            });
        });
    }

    function updateTotalFilesCard() {
        const totalFileCountCard = document.getElementById('total-file-count');
        if (totalFileCountCard) totalFileCountCard.textContent = uploadedFilesList.length;
    }

    window.deleteFile = function(index) {
        const fileName = uploadedFilesList[index].name;
        showCustomConfirm('បញ្ជាក់ការលុប', 'តើអ្នកពិតជាចង់លុបឯកសារ [' + fileName + '] នេះមែនទេ?', () => {
            uploadedFilesList.splice(index, 1);
            localStorage.setItem('uploadedExcelFiles', JSON.stringify(uploadedFilesList));
            renderTable();
            updateTotalFilesCard();
            showToast('លុបជោគជ័យ', 'ឯកសារ [' + fileName + '] ត្រូវបានលុបចេញពីប្រព័ន្ធ', 'fa-solid fa-trash', '#dc2626');
        });
    };

    function downloadFile(fileName) {
        showToast('កំពុងទាញយក...', fileName, 'fa-solid fa-circle-down', '#2563eb');
        const blob = new Blob(["នេះជាទិន្នន័យគំរូរបស់ឯកសារ៖ " + fileName], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }

    // --- ៣. មុខងារ Master Sect, Section, Style, Rampup, Prep, Process ---
    window.renderMasterSectList = function(filterText = '') {
        const container = document.getElementById('section-list-container');
        if (!container) return;
        const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
        
        if (uploadedFiles.length === 0) {
            container.innerHTML = `<div class="empty-observe-state"><p><i class="fa-solid fa-triangle-exclamation" style="color: #f59e0b;"></i> សូម Upload ហ្វាល Excel ជាមុនសិន!</p></div>`;
            return;
        }

        const allTitlesSet = new Set();
        uploadedFiles.forEach(f => f.titles && f.titles.forEach(t => allTitlesSet.add(t)));
        let titles = Array.from(allTitlesSet);
        if (filterText) titles = titles.filter(t => t.toLowerCase().includes(filterText.toLowerCase()));

        if (titles.length === 0) {
            container.innerHTML = `<div class="empty-observe-state"><p>រកមិនឃើញ Master Sect ដែលត្រូវគ្នាទេ។</p></div>`;
            return;
        }

        let html = '';
        titles.forEach(title => {
            html += `
                <div class="observe-item">
                    <div class="observe-item-left"><i class="fa-solid fa-folder-tree" style="color: #2563eb; margin-right: 8px;"></i><span class="observe-title-text">${title}</span></div>
                    <div class="observe-item-right" onclick="openTimeObserveDetail('${title}')"><i class="fa-solid fa-chevron-right"></i></div>
                </div>
            `;
        });
        container.innerHTML = html;
    };

    const searchInput = document.getElementById('section-search-input');
    if (searchInput) searchInput.addEventListener('input', (e) => renderMasterSectList(e.target.value));

    window.openTimeObserveDetail = function(title) {
        document.getElementById('time-page').style.display = 'none';
        document.getElementById('section-page').style.display = 'block';
        document.getElementById('section-page-title').innerHTML = `<i class="fa-solid fa-layer-group"></i> SECTIONS (${title})`;
        window.currentSelectedTitle = title;
        renderSectionList(title);
    };

    const backBtn = document.getElementById('back-to-observe-btn');
    if (backBtn) backBtn.addEventListener('click', () => {
        document.getElementById('section-page').style.display = 'none';
        document.getElementById('time-page').style.display = 'block';
    });

    window.renderSectionList = function(title, filterText = '') {
        const container = document.getElementById('section-item-list-container');
        if (!container) return;
        const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
        const allSectionsSet = new Set();
        uploadedFiles.forEach(f => f.sectionsMap && f.sectionsMap[title] && f.sectionsMap[title].forEach(s => allSectionsSet.add(s)));
        
        let sections = Array.from(allSectionsSet);
        sections.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

        if (filterText) sections = sections.filter(s => s.toLowerCase().includes(filterText.toLowerCase()));

        let html = '';
        sections.forEach(section => {
            html += `
                <div class="observe-item">
                    <div class="observe-item-left"><i class="fa-solid fa-folder-open" style="color: #2563eb;"></i><span class="observe-title-text">${section}</span></div>
                    <div class="observe-item-right" onclick="openStyleObservePage('${title}', '${section}')"><i class="fa-solid fa-chevron-right"></i></div>
                </div>
            `;
        });
        container.innerHTML = html || `<div class="empty-observe-state"><p>រកមិនឃើញ Section ក្រោម Title នេះទេ។</p></div>`;
    };
    

    const sectionItemSearchInput = document.getElementById('section-item-search-input');
    if (sectionItemSearchInput) sectionItemSearchInput.addEventListener('input', (e) => {
        if (window.currentSelectedTitle) renderSectionList(window.currentSelectedTitle, e.target.value);
    });

    window.openStyleObservePage = function(title, section) {
        document.getElementById('section-page').style.display = 'none';
        document.getElementById('style-page').style.display = 'block';
        document.getElementById('style-page-title').innerHTML = `<i class="fa-solid fa-shirt"></i> STYLES (${title} / ${section})`;
        window.currentSelectedTitleForStyle = title;
        window.currentSelectedSection = section;
        renderStyleList(title, section);
    };

    const backToSectionBtn = document.getElementById('back-to-section-btn');
    if (backToSectionBtn) backToSectionBtn.addEventListener('click', () => {
        document.getElementById('style-page').style.display = 'none';
        document.getElementById('section-page').style.display = 'block';
    });

    window.renderStyleList = function(title, section, filterText = '') {
        const container = document.getElementById('style-item-list-container');
        if (!container) return;
        const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
        const compositeKey = `${title}_${section}`;
        const allStylesSet = new Set();
        uploadedFiles.forEach(f => f.stylesMap && f.stylesMap[compositeKey] && f.stylesMap[compositeKey].forEach(st => allStylesSet.add(st)));

        let styles = Array.from(allStylesSet);
        styles.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
        if (filterText) styles = styles.filter(s => s.toLowerCase().includes(filterText.toLowerCase()));

        let html = '';
        styles.forEach(styleName => {
            html += `
                <div class="observe-item">
                    <div class="observe-item-left"><i class="fa-solid fa-tag" style="color: #2563eb;"></i><span class="observe-title-text">${styleName}</span></div>
                    <div class="observe-item-right" onclick="openRampupPage('${title}', '${section}', '${styleName}')"><i class="fa-solid fa-chevron-right"></i></div>
                </div>
            `;
        });
        container.innerHTML = html || `<div class="empty-observe-state"><p>រកមិនឃើញ Style ក្រោម Section នេះទេ។</p></div>`;
    };

    const styleItemSearchInput = document.getElementById('style-item-search-input');
    if (styleItemSearchInput) styleItemSearchInput.addEventListener('input', (e) => {
        if (window.currentSelectedTitleForStyle && window.currentSelectedSection) {
            renderStyleList(window.currentSelectedTitleForStyle, window.currentSelectedSection, e.target.value);
        }
    });

    window.openRampupPage = function(title, section, styleName) {
        document.getElementById('style-page').style.display = 'none';
        document.getElementById('rampup-page').style.display = 'block';
        document.getElementById('rampup-page-title').innerHTML = `<i class="fa-solid fa-calendar-day"></i> RAMPUP DAYS (${title} / ${section} / ${styleName})`;
        window.currentSelectedTitleForRampup = title;
        window.currentSelectedSectionForRampup = section;
        window.currentSelectedStyle = styleName;
        renderRampupList(title, section, styleName);
    };

    const backToStyleBtn = document.getElementById('back-to-style-btn');
    if (backToStyleBtn) backToStyleBtn.addEventListener('click', () => {
        document.getElementById('rampup-page').style.display = 'none';
        document.getElementById('style-page').style.display = 'block';
    });

    window.renderRampupList = function(title, section, styleName, filterText = '') {
        const container = document.getElementById('rampup-item-list-container');
        if (!container) return;
        const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
        const compositeKey = `${title}_${section}_${styleName}`;
        const allRampupSet = new Set();
        uploadedFiles.forEach(f => f.rampupMap && f.rampupMap[compositeKey] && f.rampupMap[compositeKey].forEach(rd => allRampupSet.add(rd)));

        let rampups = Array.from(allRampupSet);
        rampups.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
        if (filterText) rampups = rampups.filter(r => r.toLowerCase().includes(filterText.toLowerCase()));

        let html = '';
        rampups.forEach(rampupDay => {
            html += `
                <div class="observe-item">
                    <div class="observe-item-left"><i class="fa-solid fa-calendar-check" style="color: #2563eb;"></i><span class="observe-title-text">${rampupDay}</span></div>
                    <div class="observe-item-right" onclick="openPrepAssemblyPage('${title}', '${section}', '${styleName}', '${rampupDay}')"><i class="fa-solid fa-chevron-right"></i></div>
                </div>
            `;
        });
        container.innerHTML = html || `<div class="empty-observe-state"><p>រកមិនឃើញ RampupDay ក្រោម Style នេះទេ។</p></div>`;
    };

    const rampupItemSearchInput = document.getElementById('rampup-item-search-input');
    if (rampupItemSearchInput) rampupItemSearchInput.addEventListener('input', (e) => {
        if (window.currentSelectedTitleForRampup && window.currentSelectedSectionForRampup && window.currentSelectedStyle) {
            renderRampupList(window.currentSelectedTitleForRampup, window.currentSelectedSectionForRampup, window.currentSelectedStyle, e.target.value);
        }
    });

    window.openPrepAssemblyPage = function(title, section, styleName, rampupDay) {
        document.getElementById('rampup-page').style.display = 'none';
        document.getElementById('prep-page').style.display = 'block';
        document.getElementById('prep-page-title').innerHTML = `<i class="fa-solid fa-list-check"></i> ${rampupDay}`;
        window.prepContext = { title, section, styleName, rampupDay };
        renderPrepAssemblyList(title, section, styleName, rampupDay);
    };

    const backToRampupBtn = document.getElementById('back-to-rampup-btn');
    if (backToRampupBtn) backToRampupBtn.addEventListener('click', () => {
        document.getElementById('prep-page').style.display = 'none';
        document.getElementById('rampup-page').style.display = 'block';
    });

    // 🌟 មុខងារ renderPrepAssemblyList កែសម្រួលថ្មី (ពិនិត្យចំនួន Process ឱ្យបានគ្រប់គ្រាន់ពិតប្រាកដ)
    window.renderPrepAssemblyList = function(title, section, styleName, rampupDay, filterText = '') {
        const container = document.getElementById('prep-item-list-container');
        if (!container) return;
        const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
        const savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];
        const compositeKey = `${title}_${section}_${styleName}_${rampupDay}`;
        let rawItems = [];
        uploadedFiles.forEach(f => f.prepMap && f.prepMap[compositeKey] && (rawItems = rawItems.concat(f.prepMap[compositeKey])));

        const groupedMap = {};
        rawItems.forEach(item => {
            const prepName = item.prepAssembly;
            if (!groupedMap[prepName]) {
                groupedMap[prepName] = { 
                    prepAssembly: prepName, 
                    primaryKeys: new Set(), 
                    smvsList: [],
                    totalProcesses: 0,           // ចំនួន Process សរុបពិតប្រាកដក្នុង Group នេះ
                    matchedProcesses: new Set()  // រាប់ចំនួន Process ជាក់លាក់ដែលបាន Save រួច
                };
            }
            if (item.primaryKey) groupedMap[prepName].primaryKeys.add(item.primaryKey);
            
            groupedMap[prepName].totalProcesses++;

            // 🌟 ពិនិត្យផ្ទៀងផ្ទាត់យ៉ាងម៉ត់ចត់៖ ត្រូវតែចំ Style, RampupDay, PrepAssembly និង Process Name ជាក់លាក់
            const isStudied = savedStudies.some(s => {
                const sProcess = (s.processRegion || s.processName || '').trim().toLowerCase();
                const iProcess = (item.processRegion || item.processName || '').trim().toLowerCase();
                const sPrep = (s.prepAssembly || s.prepName || '').trim();
                const iPrep = prepName.trim();

                return s.styleName === styleName && 
                       s.rampupDay === rampupDay && 
                       sPrep === iPrep && 
                       sProcess === iProcess;
            });

            if (isStudied) {
                groupedMap[prepName].matchedProcesses.add(item.processRegion || item.processName || item.primaryKey);
            }

            const numSmv = parseFloat(item.smv);
            if (!isNaN(numSmv)) {
                groupedMap[prepName].smvsList.push(numSmv);
            }
        });

        let prepItems = Object.values(groupedMap).map(g => {
            const totalSmv = g.smvsList.reduce((acc, curr) => acc + curr, 0);
            
            // 🌟 ធានាថាចំនួន Process ដែលបាន Save ត្រូវតែស្មើ ឬច្រើនជាងចំនួនសរុប ទើបถือว่า Completed
            const completedCount = g.matchedProcesses.size;
            const isCompleted = g.totalProcesses > 0 && completedCount >= g.totalProcesses;

            return {
                prepAssembly: g.prepAssembly,
                primaryKey: Array.from(g.primaryKeys).join(', '),
                smv: totalSmv > 0 ? totalSmv.toFixed(2) : '0.00',
                totalProcesses: g.totalProcesses,
                completedCount: completedCount,
                isCompleted: isCompleted 
            };
        });

        let html = '';
        prepItems.forEach(item => {
            let statusBadgeHtml = '';

            if (item.isCompleted) {
                // 🟢 បង្ហាញសញ្ញាគ្រីសបៃតង ពេលចាប់គ្រប់ចំនួន
                statusBadgeHtml = `
                    <span style="background: #d1fae5; color: #047857; padding: 4px 12px; border-radius: 20px; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                        <i class="fa-solid fa-circle-check"></i> សូមអបអរសាទរ👏🎉 អ្នកបានចាប់នាទីរួចរាល់ (${item.completedCount}/${item.totalProcesses})
                    </span>
                `;
            } else {
                // 🔴 បង្ហាញសញ្ញាក្រហម ពេលចាប់មិនទាន់គ្រប់ (ព្រមទាំងបង្ហាញចំនួនសម្រេចបាន)
                statusBadgeHtml = `
                    <span style="background: #fee2e2; color: #dc2626; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: bold; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                        <i class="fa-solid fa-triangle-exclamation"></i> បញ្ជាក់៖ អ្នកមិនទាន់ចាប់នាទីគ្រប់ចំនួនProcessនោះទេ😁😁!! (${item.completedCount}/${item.totalProcesses})
                    </span>
                `;
            }

            html += `
                <div class="observe-item">
                    <div class="observe-item-left" style="flex-direction: column; align-items: flex-start; gap: 8px; width: 100%;">
                        <div style="display: flex; align-items: center; width: 100%; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
                            <span class="observe-title-text" style="font-size: 16px; font-weight: bold;">${item.prepAssembly}</span>
                            <div>${statusBadgeHtml}</div>
                        </div>
                        
                        <div style="display: flex; gap: 15px; font-size: 12px; color: #94a3b8; align-items: center; width: 100%; justify-content: space-between; flex-wrap: wrap;">
                            <span>${item.primaryKey}</span>
                            <span style="font-weight: bold; color: #0284c7; background: #e0f2fe; padding: 3px 10px; border-radius: 6px; font-size: 13px;">
                                Total SMV: ${item.smv}
                            </span>
                        </div>
                    </div>
                    <div class="observe-item-right" onclick="openProcessPage('${title}', '${section}', '${styleName}', '${rampupDay}', '${item.prepAssembly}')">
                        <i class="fa-solid fa-chevron-right"></i>
                    </div>
                </div>
            `;
        });
        container.innerHTML = html || `<div class="empty-observe-state"><p>រកមិនឃើញទិន្នន័យ Prep / Assembly ក្រោម RampupDay នេះទេ។</p></div>`;
    };

    

    const prepItemSearchInput = document.getElementById('prep-item-search-input');
    if (prepItemSearchInput) prepItemSearchInput.addEventListener('input', (e) => {
        if (window.prepContext) renderPrepAssemblyList(window.prepContext.title, window.prepContext.section, window.prepContext.styleName, window.prepContext.rampupDay, e.target.value);
    });

    window.openProcessPage = function(title, section, styleName, rampupDay, prepName) {
        document.getElementById('prep-page').style.display = 'none';
        document.getElementById('process-page').style.display = 'block';
        document.getElementById('process-page-title').innerHTML = `<i class="fa-solid fa-list-check"></i> ${prepName}`;
        window.processContext = { title, section, styleName, rampupDay, prepName };
        renderProcessList(title, section, styleName, rampupDay, prepName);
    };

    const backToPrepBtn = document.getElementById('back-to-prep-btn');
    if (backToPrepBtn) backToPrepBtn.addEventListener('click', () => {
        document.getElementById('process-page').style.display = 'none';
        document.getElementById('prep-page').style.display = 'block';
    });

// 🌟 មុខងារ renderProcessList (រួមបញ្ចូលទាំង CT, PFM%, និងប៊ូតុង Refresh សមរម្យ)
    window.renderProcessList = function(title, section, styleName, rampupDay, prepName, filterText = '') {
        const container = document.getElementById('process-item-list-container');
        if (!container) return;
        const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
        const savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];
        const compositeKey = `${title}_${section}_${styleName}_${rampupDay}`;
        let rawItems = [];
        uploadedFiles.forEach(f => f.prepMap && f.prepMap[compositeKey] && (rawItems = rawItems.concat(f.prepMap[compositeKey])));

        let processItems = rawItems.filter(item => item.prepAssembly === prepName);
        if (filterText) {
            const kw = filterText.toLowerCase();
            processItems = processItems.filter(item => (item.processRegion && item.processRegion.toLowerCase().includes(kw)) || (item.primaryKey && item.primaryKey.toLowerCase().includes(kw)));
        }

        if (processItems.length === 0) {
            container.innerHTML = `<div class="empty-observe-state"><p>រកមិនឃើញ Process ក្រោម Prep នេះទេ។</p></div>`;
            return;
        }

        let html = '';
        processItems.forEach((item, index) => {
            const khmerText = item.processRegion || 'អត់មានឈ្មោះខ្មែរ';
            const thaiText = item.processName || '';
            const sewerNameText = (item.sewerName && item.sewerName !== 'undefined' && item.sewerName.trim() !== '') ? item.sewerName : 'No Sewer';
            const primaryKeyText = item.primaryKey || 'N/A';
            const processNo = item.no || (index + 1);
            
            let numSmv = parseFloat(item.smv);
            const smvText = !isNaN(numSmv) ? numSmv.toFixed(2) : (item.smv || '0.00');

            const matchedStudy = savedStudies.find(s => {
                const sProcess = (s.processRegion || s.processName || '').trim().toLowerCase();
                const iProcess = (item.processRegion || item.processName || '').trim().toLowerCase();
                const sSewer = (s.sewerName || '').trim().toLowerCase();
                const iSewer = sewerNameText.trim().toLowerCase();
                const sStyle = (s.styleName || '').trim();
                const iStyle = styleName.trim();

                return sStyle === iStyle && sProcess === iProcess && sSewer === iSewer;
            });

            let checkmarkBadgeHtml = '';
            let avgCtBadgeHtml = '';
            let pfmBadgeHtml = '';
            let deleteBtnHtml = '';

            if (matchedStudy) {
                checkmarkBadgeHtml = `
                    <span style="background: #d1fae5; color: #047857; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; display: inline-flex; align-items: center; gap: 4px;">
                        <i class="fa-solid fa-circle-check"></i> រួចរាល់
                    </span>
                `;

                const avgCtVal = matchedStudy.averageCT || '0.00s';
                avgCtBadgeHtml = `
                    <span style="background: #e0f2fe; color: #0369a1; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 5px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                        <i class="fa-solid fa-stopwatch"></i> CT: ${avgCtVal}
                    </span>
                `;

                const pfmVal = parseFloat(matchedStudy.pfm) || 0;
                let pfmBg = '#dbeafe';
                let pfmColor = '#1d4ed8';
                if (pfmVal < 84) { pfmBg = '#fee2e2'; pfmColor = '#dc2626'; }
                else if (pfmVal >= 85 && pfmVal <= 99) { pfmBg = '#fef3c7'; pfmColor = '#d97706'; }
                else if (pfmVal >= 100) { pfmBg = '#e0f2fe'; pfmColor = '#0284c7'; }

                pfmBadgeHtml = `
                    <span style="background: ${pfmBg}; color: ${pfmColor}; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                        <i class="fa-solid fa-chart-pie"></i> PFM: ${pfmVal}%
                    </span>
                `;

                // 🔄 ប៊ូតុង Refresh (Reset ទិន្នន័យ Time Study)
                deleteBtnHtml = `
                    <button type="button" class="delete-study-btn" data-study-id="${matchedStudy.id}" title="Reset / ចាប់ផ្ដើមការចាប់នាទីសាថ្មី" style="background: #fee2e2; color: #dc2626; border: none; padding: 4px 10px; border-radius: 6px; font-size: 13px; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; transition: all 0.2s;">
                        <i class="fa-solid fa-arrows-rotate"></i> Refresh
                    </button>
                `;
            }

            html += `
                <div class="observe-item" style="display: flex; align-items: center; justify-content: space-between; padding: 14px 18px; margin-bottom: 12px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; transition: all 0.2s ease;">
                    
                    <div class="observe-item-left process-item-btn" data-index="${index}" style="flex-direction: column; align-items: flex-start; gap: 8px; width: 100%; cursor: pointer;">
                        
                        <div style="font-size: 15px; font-weight: 600; color: #1e293b; display: flex; align-items: center; width: 100%; justify-content: space-between;">
                            <div>
                                <span style="color: #2563eb; margin-right: 5px;">${processNo}.</span> ${khmerText}
                            </div>
                            <div style="display: flex; align-items: center; gap: 8px;">
                                ${deleteBtnHtml}
                                ${checkmarkBadgeHtml}
                            </div>
                        </div>

                        <span style="font-size: 13px; color: #64748b;">${thaiText}</span>
                        
                        <div style="display: flex; gap: 15px; font-size: 13px; color: #475569; background: #f8fafc; padding: 8px 12px; border-radius: 8px; width: 100%; flex-wrap: wrap; align-items: center; border: 1px solid #f1f5f9;">
                            <span><b>Seq:</b> ${item.seq || 99}</span>
                            <span><b>Sewer:</b> ${sewerNameText}</span>
                            <span><b>MC:</b> ${item.mc || '-'}</span>
                        </div>

                        <div style="display: flex; width: 100%; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-top: 2px;">
                            <span style="font-size: 12px; color: #64748b;"><b>PrimaryKey:</b> ${primaryKeyText}</span>
                            <div style="display: flex; align-items: center; gap: 8px;">
                                ${pfmBadgeHtml}
                                ${avgCtBadgeHtml}
                                <span style="font-weight: bold; color: #584ff4; background: #6661cb5d; padding: 4px 10px; border-radius: 6px; border: 1px solid #fefefe; font-size: 13px;">
                                    SMV: ${smvText}
                                </span>
                            </div>
                        </div>

                    </div>

                    <div class="observe-item-right process-item-btn" data-index="${index}" style="margin-left: 15px; color: #94a3b8; cursor: pointer;">
                        <i class="fa-solid fa-chevron-right"></i>
                    </div>
                </div>
            `;
        });
        container.innerHTML = html;

        // 🔗 Event សម្រាប់ចុចចូលទៅកាន់ Time Study Form
        const itemButtons = container.querySelectorAll('.process-item-btn');
        itemButtons.forEach(btn => {
            btn.addEventListener('click', function() {
                const idx = parseInt(this.getAttribute('data-index'));
                const selectedItem = processItems[idx];
                openTimeStudyForm(selectedItem);
            });
        });

// 🔄 Event សម្រាប់ចុចប៊ូតុង Refresh (បង្រួមរូបសញ្ញា និងកែ Border Radius ឱ្យមូលស្អាត)
        const deleteButtons = container.querySelectorAll('.delete-study-btn');
        deleteButtons.forEach(delBtn => {
            delBtn.addEventListener('click', function(e) {
                e.stopPropagation(); // មិនឱ្យរត់ចូល Form ពេលចុចប៊ូតុង Refresh
                const studyId = this.getAttribute('data-study-id');
                
                Swal.fire({
                    title: '<span style="font-size: 17px; font-weight: bold;">តើអ្នកពិតជាចង់ Reset?</span>',
                    html: '<span style="font-size: 13px; color: #64748b;">ឬចាប់ផ្ដើមការចាប់នាទី Process នេះសាថ្មីមែនទេ?</span>',
                    icon: 'warning',
                    width: '300px',
                    padding: '1.2rem',
                    showCancelButton: true,
                    confirmButtonColor: '#d33',
                    cancelButtonColor: '#3085d6',
                    confirmButtonText: '<span style="font-size: 13px;">បាទ/ចាស, Reset!</span>',
                    cancelButtonText: '<span style="font-size: 13px;">បោះបង់</span>',
                    // 🌟 កែសម្រួល Border Radius និងទំហំរូបសញ្ញា (icon) តាមរយៈ customClass
                    customClass: {
                        popup: 'custom-swal-popup',
                        icon: 'custom-swal-icon'
                    }
                }).then((result) => {
                    if (result.isConfirmed) {
                        let savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];
                        savedStudies = savedStudies.filter(s => s.id !== studyId);
                        localStorage.setItem('savedTimeStudies', JSON.stringify(savedStudies));

                        renderProcessList(title, section, styleName, rampupDay, prepName, filterText);
                        
                        Swal.fire({
                            title: '<span style="font-size: 15px;">បានជោគជ័យ!</span>',
                            html: '<span style="font-size: 12px;">ទិន្នន័យត្រូវបាន Reset សាថ្មីរួចរាល់។</span>',
                            icon: 'success',
                            width: '300px',
                            timer: 1500,
                            showConfirmButton: false,
                            customClass: {
                                popup: 'custom-swal-popup'
                            }
                        });
                    }
                });
            });
        });
    };

    const processItemSearchInput = document.getElementById('process-item-search-input');
    if (processItemSearchInput) processItemSearchInput.addEventListener('input', (e) => {
        if (window.processContext) renderProcessList(window.processContext.title, window.processContext.section, window.processContext.styleName, window.processContext.rampupDay, window.processContext.prepName, e.target.value);
    });

    // --- ៤. មុខងារ Time Study Form, Stopwatch & Edit Modal System ---
    let stopwatchInterval = null;
    let stopwatchTime = 0;
    let stopwatchState = 'STOPPED'; 
    let lapRecords = [];
    let lastLapTotalTime = 0; 
    let currentStartTimeStamp = '';

    window.openTimeStudyForm = function(itemData) {
        document.getElementById('process-page').style.display = 'none';
        document.getElementById('timestudy-page').style.display = 'block';

        const validSewer = (itemData.sewerName && itemData.sewerName !== 'undefined' && itemData.sewerName.trim() !== '') ? itemData.sewerName : 'No Sewer';
        
        let numSmv = parseFloat(itemData.smv);
        let formattedSmv = !isNaN(numSmv) ? numSmv.toFixed(2) : (itemData.smv || '0.00');

        document.getElementById('ts-process-khmer').textContent = itemData.processRegion || itemData.processName || 'N/A';
        document.getElementById('ts-process-thai').textContent = itemData.processName || '';
        document.getElementById('ts-primary-key').textContent = itemData.primaryKey || 'N/A';
        document.getElementById('ts-smv').textContent = formattedSmv;

        const sewerBadge = document.getElementById('ts-sewer-badge');
        if (sewerBadge) {
            sewerBadge.textContent = validSewer;
        }

        window.timeStudyContext = { ...itemData, smv: formattedSmv, sewerName: validSewer };
        resetStopwatch();
    };

    const backToProcessBtn = document.getElementById('back-to-process-btn');
    if (backToProcessBtn) backToProcessBtn.addEventListener('click', () => {
        clearInterval(stopwatchInterval);
        document.getElementById('timestudy-page').style.display = 'none';
        document.getElementById('process-page').style.display = 'block';
    });

    const mainActionBtn = document.getElementById('ts-main-action-btn'); 
    const secondaryActionBtn = document.getElementById('ts-lap-btn');    
    const display = document.getElementById('stopwatch-display');

    if (mainActionBtn) {
        mainActionBtn.addEventListener('click', () => {
            if (stopwatchState === 'STOPPED' || stopwatchState === 'PAUSED') {
                startStopwatch();
            } else if (stopwatchState === 'RUNNING') {
                pauseStopwatch();
            }
        });
    }

    if (secondaryActionBtn) {
        secondaryActionBtn.addEventListener('click', () => {
            if (stopwatchState === 'RUNNING') {
                recordLap();
            } else if (stopwatchState === 'PAUSED') {
                resetStopwatch();
            }
        });
    }

    function startStopwatch() {
        if (!currentStartTimeStamp || stopwatchState === 'STOPPED') {
            const now = new Date();
            currentStartTimeStamp = now.toLocaleDateString('km-KH') + ' ' + now.toLocaleTimeString('km-KH', {hour: '2-digit', minute:'2-digit', second:'2-digit'});
        }

        stopwatchState = 'RUNNING';
        mainActionBtn.innerHTML = `<i class="fa-solid fa-pause"></i> Pause`;
        mainActionBtn.style.backgroundColor = '#f59e0b';
        secondaryActionBtn.innerHTML = `<i class="fa-solid fa-flag"></i> កត់ត្រា (Lap)`;
        secondaryActionBtn.style.backgroundColor = '#e2e8f0';
        secondaryActionBtn.style.color = '#1e293b';
        secondaryActionBtn.removeAttribute('disabled');

        const startTime = Date.now() - stopwatchTime;
        stopwatchInterval = setInterval(() => {
            stopwatchTime = Date.now() - startTime;
            updateDisplay();
        }, 10);
    }

    function pauseStopwatch() {
        stopwatchState = 'PAUSED';
        mainActionBtn.innerHTML = `<i class="fa-solid fa-play"></i> Play`;
        mainActionBtn.style.backgroundColor = '#10b981';
        secondaryActionBtn.innerHTML = `<i class="fa-solid fa-rotate-left"></i> Restart`;
        secondaryActionBtn.style.backgroundColor = '#fee2e2';
        secondaryActionBtn.style.color = '#dc2626';
        clearInterval(stopwatchInterval);
    }

    function resetStopwatch() {
        clearInterval(stopwatchInterval);
        stopwatchState = 'STOPPED';
        stopwatchTime = 0;
        lastLapTotalTime = 0; 
        currentStartTimeStamp = '';
        lapRecords = [];
        updateDisplay();
        renderLapTable();

        if (mainActionBtn) {
            mainActionBtn.innerHTML = `<i class="fa-solid fa-play"></i> Play`;
            mainActionBtn.style.backgroundColor = '#10b981';
        }
        if (secondaryActionBtn) {
            secondaryActionBtn.innerHTML = `<i class="fa-solid fa-flag"></i> កត់ត្រា (Lap)`;
            secondaryActionBtn.style.backgroundColor = '#e2e8f0';
            secondaryActionBtn.style.color = '#94a3b8';
            secondaryActionBtn.setAttribute('disabled', 'true');
        }
    }

    function updateDisplay() {
        if (!display) return;
        const minutes = Math.floor(stopwatchTime / 60000);
        const seconds = Math.floor((stopwatchTime % 60000) / 1000);
        const milliseconds = Math.floor((stopwatchTime % 1000) / 10);

        display.textContent = 
            String(minutes).padStart(2, '0') + ':' + 
            String(seconds).padStart(2, '0') + '.' + 
            String(milliseconds).padStart(2, '0');
    }

    function recordLap() {
        const currentTotalTime = stopwatchTime; 
        const lapDuration = currentTotalTime - lastLapTotalTime;
        lastLapTotalTime = currentTotalTime; 

        const minutes = Math.floor(lapDuration / 60000);
        const seconds = Math.floor((lapDuration % 60000) / 1000);
        const milliseconds = Math.floor((lapDuration % 1000) / 10);

        const formattedLapTime = 
            String(minutes).padStart(2, '0') + ':' + 
            String(seconds).padStart(2, '0') + '.' + 
            String(milliseconds).padStart(2, '0');

        lapRecords.push({
            round: lapRecords.length + 1,
            time: formattedLapTime,
            durationMs: lapDuration
        });
        renderLapTable();
    }

    function renderLapTable() {
        const tbody = document.getElementById('lap-table-body');
        if (!tbody) return;

        if (lapRecords.length === 0) {
            tbody.innerHTML = `<tr><td colspan="3" style="text-align: center; color: #94a3b8;">មិនទាន់មានទិន្នន័យចាប់ម៉ោងនៅឡើយទេ</td></tr>`;
            return;
        }

        let html = '';
        lapRecords.forEach((lap, index) => {
            html += `
                <tr>
                    <td>Round ${lap.round}</td>
                    <td style="font-weight: 600; color: #2563eb;">${lap.time}</td>
                    <td>
                        <button onclick="deleteLap(${index})" style="background: none; border: none; color: #dc2626; cursor: pointer;"><i class="fa-solid fa-trash"></i></button>
                    </td>
                </tr>
            `;
        });
        tbody.innerHTML = html;
    }

    window.deleteLap = function(index) {
        lapRecords.splice(index, 1);
        lapRecords.forEach((lap, i) => lap.round = i + 1);
        renderLapTable();
    };

// --- 5. មុខងារ Save Time Study (កែសម្រួលថ្មី ឲ្យ Refresh ទាំង Process និង Prep Page ភ្លាមៗ) ---
    const saveBtn = document.getElementById('ts-save-btn');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            if (lapRecords.length < 3) {
                showCustomAlert('មិនអាចរក្សាទុកបានទេ', 'សូមធ្វើការកត់ត្រា Lap យ៉ាងហោចណាស់ ៣ ដង (Round) ជាមុនសិន!', 'fa-solid fa-triangle-exclamation', '#f59e0b');
                return;
            }

            let totalMs = 0;
            lapRecords.forEach(lap => {
                totalMs += lap.durationMs;
            });
            const avgMs = totalMs / lapRecords.length;
            const totalSeconds = avgMs / 1000; 

            const secPart = Math.floor(totalSeconds);
            const msPart = Math.floor((avgMs % 1000) / 10);
            const avgFormattedTime = String(secPart).padStart(2, '0') + '.' + String(msPart).padStart(2, '0') + 's';

            const context = window.timeStudyContext || {};
            const smvVal = parseFloat(context.smv) || 0; 
            
            let pfmPercent = 0;
            if (totalSeconds > 0 && smvVal > 0) {
                pfmPercent = (smvVal / totalSeconds) * 100;
            }

            let outputPcs = 0;
            if (totalSeconds > 0) {
                outputPcs = Math.round(3600 / totalSeconds);
            }

            const now = new Date();
            const dateTimeEndVal = now.toLocaleDateString('km-KH') + ' ' + now.toLocaleTimeString('km-KH', {hour: '2-digit', minute:'2-digit', second:'2-digit'});
            const startTimeVal = currentStartTimeStamp || dateTimeEndVal;

            const studyRecord = {
                id: 'TS_' + Date.now(),
                date: dateTimeEndVal,
                dateTimeStart: startTimeVal,
                dateTimeEnd: dateTimeEndVal,
                title: context.title || 'N/A',
                section: context.section || 'N/A',
                styleName: context.styleName || 'N/A',
                rampupDay: context.rampupDay || 'N/A',
                prepName: context.prepName || context.prepAssembly || 'N/A',
                prepAssembly: context.prepAssembly || context.prepName || 'N/A',
                sewerName: context.sewerName || 'No Sewer',
                processRegion: context.processRegion || 'N/A',
                processName: context.processName || 'N/A',
                primaryKey: context.primaryKey || 'N/A',
                smv: context.smv || '0.00',
                no: context.no || '-',
                seq: parseInt(context.seq) || 99,
                operatorSeq: parseInt(context.seq) || 99,
                mc: context.mc || '-',
                weight: context.weight || '-',
                laps: [...lapRecords],
                averageCT: avgFormattedTime,
                avgSeconds: totalSeconds, 
                output: outputPcs,
                pfm: pfmPercent.toFixed(1)
            };

            let savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];
            savedStudies.unshift(studyRecord);
            localStorage.setItem('savedTimeStudies', JSON.stringify(savedStudies));

            // 🌟 ជូនដំណឹងជោគជ័យ
            showToast('ជោគជ័យ!', 'ទិន្នន័យ Time Study ត្រូវបានរក្សាទុក', 'fa-solid fa-circle-check', '#10b981');
            resetStopwatch();

            // 🌟 ធ្វើការ Refresh ទិន្នន័យទាំង Prep Page និង Process Page ភ្លាមៗមុននឹងប្ដូរទំព័របង្ហាញ
            if (window.prepContext) {
                renderPrepAssemblyList(
                    window.prepContext.title, 
                    window.prepContext.section, 
                    window.prepContext.styleName, 
                    window.prepContext.rampupDay
                );
            }

            if (window.processContext) {
                renderProcessList(
                    window.processContext.title, 
                    window.processContext.section, 
                    window.processContext.styleName, 
                    window.processContext.rampupDay, 
                    window.processContext.prepName
                );
            }

            // 🌟 ប្ដូរទំព័រត្រឡប់មកកាន់ Process Page វិញ
            document.getElementById('timestudy-page').style.display = 'none';
            document.getElementById('process-page').style.display = 'block';
        });
    }
    // --- ៥.១ Dynamic Edit Modal System ---
    window.openTimeObserveEditModal = function() {
        let studyData = window.timeStudyContext;

        if (!studyData || !studyData.primaryKey) {
            const pkElement = document.getElementById('ts-primary-key');
            const pkText = pkElement ? pkElement.textContent.trim() : '';
            
            const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
            for (let f of uploadedFiles) {
                if (f.prepMap) {
                    for (let key in f.prepMap) {
                        const found = f.prepMap[key].find(item => item.primaryKey === pkText);
                        if (found) { studyData = found; break; }
                    }
                }
                if (studyData) break;
            }
        }

        if (!studyData) {
            const pkElement = document.getElementById('ts-primary-key');
            const pkText = pkElement ? pkElement.textContent.trim() : '';
            const savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];
            studyData = savedStudies.find(item => item.primaryKey === pkText);
        }

        if (!studyData) {
            const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
            for (let f of uploadedFiles) {
                if (f.prepMap) {
                    for (let key in f.prepMap) {
                        if (f.prepMap[key].length > 0) {
                            studyData = f.prepMap[key][0];
                            break;
                        }
                    }
                }
                if (studyData) break;
            }
        }

        if (!studyData) {
            alert('សូមជ្រើសរើស Process ណាមួយជាមុនសិន!');
            return;
        }

        window.timeStudyContext = studyData;

        const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];
        const fileObj = uploadedFiles[0]; 
        const excelHeaders = fileObj && fileObj.excelHeaders ? fileObj.excelHeaders : ['Title', 'PrimaryKey', 'RampupDay', 'Section', 'Prep_Assembly', 'DateTime-Start', 'No', 'Seq', 'ProcessRunNo', 'SewerName', 'MC', 'Weight', 'SMV'];

        const rawRow = studyData.rawRowData || {};

        const container = document.getElementById('dynamic-edit-inputs-container');
        if (container) {
            let html = '';
            excelHeaders.forEach(header => {
                let val = '';
                const cleanHeader = header.trim().toLowerCase().replace(/[\s_]+/g, '');

                const matchedKey = Object.keys(rawRow).find(k => k.trim().toLowerCase().replace(/[\s_]+/g, '') === cleanHeader);
                
                if (matchedKey && rawRow[matchedKey] !== undefined && rawRow[matchedKey] !== null && String(rawRow[matchedKey]).trim() !== '') {
                    val = rawRow[matchedKey];
                } else {
                    if (cleanHeader === 'no') val = studyData.no || '';
                    else if (cleanHeader === 'seq' || cleanHeader === 'sequence') val = studyData.seq || studyData.operatorSeq || '';
                    else if (cleanHeader === 'mc' || cleanHeader === 'machine') val = studyData.mc || '';
                    else if (cleanHeader === 'weight' || cleanHeader === 'wt') val = studyData.weight || '';
                    else if (cleanHeader.includes('datetime') || cleanHeader.includes('start')) val = studyData.dateTimeStart || '';
                    else if (cleanHeader.includes('title') || cleanHeader.includes('factory')) val = studyData.title || studyData.factory || '';
                    else if (cleanHeader.includes('section')) val = studyData.section || '';
                    else if (cleanHeader.includes('style')) val = studyData.styleName || '';
                    else if (cleanHeader.includes('rampup')) val = studyData.rampupDay || '';
                    else if (cleanHeader.includes('prep') || cleanHeader.includes('assembly')) val = studyData.prepAssembly || studyData.prepName || '';
                    else if (cleanHeader.includes('sewer')) val = studyData.sewerName || studyData.sewer || '';
                    else if (cleanHeader.includes('process')) val = studyData.processRegion || studyData.processName || '';
                    else if (cleanHeader.includes('smv')) val = studyData.smv || '';
                    else if (cleanHeader.includes('primary')) val = studyData.primaryKey || '';
                }

                html += `
                    <div style="display: flex; flex-direction: column; gap: 5px;">
                        <label style="font-size: 12px; color: #334155; font-weight: bold;">${header}</label>
                        <input type="text" class="dynamic-obs-input" data-header="${header}" value="${val !== undefined && val !== null ? val : ''}" style="padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 13px; outline: none; width: 100%;">
                    </div>
                `;
            });
            container.innerHTML = html;
        }

        document.getElementById('timeObserveEditModal').style.display = 'flex';
    };

    window.closeTimeObserveEditModal = function() {
        document.getElementById('timeObserveEditModal').style.display = 'none';
    };

    window.saveTimeObserveEdit = function() {
        const inputs = document.querySelectorAll('.dynamic-obs-input');
        if (!window.timeStudyContext) window.timeStudyContext = {};

        if (!window.timeStudyContext.rawRowData) {
            window.timeStudyContext.rawRowData = {};
        }

        inputs.forEach(input => {
            const header = input.getAttribute('data-header');
            const val = input.value;
            window.timeStudyContext.rawRowData[header] = val;

            const lowerH = header.toLowerCase();
            if (lowerH === 'no') window.timeStudyContext.no = val;
            else if (lowerH === 'seq') { window.timeStudyContext.seq = val; window.timeStudyContext.operatorSeq = parseInt(val) || 99; }
            else if (lowerH === 'mc' || lowerH === 'machine') window.timeStudyContext.mc = val;
            else if (lowerH === 'weight' || lowerH === 'wt') window.timeStudyContext.weight = val;
            else if (lowerH.includes('datetime') || lowerH.includes('start')) window.timeStudyContext.dateTimeStart = val;
            else if (lowerH.includes('title') || lowerH.includes('factory')) { window.timeStudyContext.title = val; window.timeStudyContext.factory = val; }
            else if (lowerH.includes('section')) window.timeStudyContext.section = val;
            else if (lowerH.includes('style')) window.timeStudyContext.styleName = val;
            else if (lowerH.includes('prep') || lowerH.includes('assembly')) { window.timeStudyContext.prepAssembly = val; window.timeStudyContext.prepName = val; }
            else if (lowerH.includes('sewer')) { window.timeStudyContext.sewerName = val; window.timeStudyContext.sewer = val; }
            else if (lowerH.includes('process')) window.timeStudyContext.processRegion = val;
            else if (lowerH.includes('smv')) window.timeStudyContext.smv = val;
        });

        document.getElementById('ts-process-khmer').textContent = window.timeStudyContext.processRegion || 'N/A';
        const sewerBadge = document.getElementById('ts-sewer-badge');
        if (sewerBadge) {
            sewerBadge.textContent = window.timeStudyContext.sewerName || 'No Sewer';
        }

        closeTimeObserveEditModal();

        if (typeof showToast === 'function') {
            showToast('ជោគជ័យ', 'រាល់ទិន្នន័យក្បាលតារាងទាំងអស់ត្រូវបានកែសម្រួលរួចរាល់!', 'fa-solid fa-check', '#10b981');
        } else {
            alert('កែសម្រួលទិន្នន័យជោគជ័យ!');
        }
    };

    // --- ៦. មុខងារ Dashboard & Graph (Sorted by Seq) ---
    window.renderDashboardStats = function() {
        const dashboardPage = document.getElementById('dashboard-page');
        if (!dashboardPage) return;

        const savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];
        const uploadedFiles = JSON.parse(localStorage.getItem('uploadedExcelFiles')) || [];

        const factories = [...new Set(savedStudies.map(item => item.factory || item.title).filter(Boolean))];
        const sections = [...new Set(savedStudies.map(item => item.section).filter(Boolean))];
        const styles = [...new Set(savedStudies.map(item => item.styleName).filter(Boolean))];
        const rampupDays = [...new Set(savedStudies.map(item => item.rampupDay).filter(Boolean))];

        dashboardPage.innerHTML = `
            <div class="observe-container">
                <div class="observe-header">
                    <h2><i class="fa-solid fa-chart-line"></i> DASHBOARD ANALYSIS & TIME STUDIES</h2>
                </div>

                <div class="cards-grid" style="margin-top: 20px;">
                    <div class="card">
                        <div class="card-icon card-1 icon"><i class="fa-solid fa-file-excel"></i></div>
                        <div class="card-info">
                            <h3>Total Excel Files</h3>
                            <p id="total-file-count">${uploadedFiles.length}</p>
                        </div>
                    </div>
                    <div class="card">
                        <div class="card-icon card-2 icon"><i class="fa-solid fa-stopwatch"></i></div>
                        <div class="card-info">
                            <h3>Total Studies Saved</h3>
                            <p>${savedStudies.length}</p>
                        </div>
                    </div>
                </div>

                <!-- Filter Selection Bar -->
                <div style="background: #e0f2fe; padding: 15px; border-radius: 10px; display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 10px; margin-top: 25px; border: 1px solid #bae6fd;">
                    <div style="background: #fff; padding: 8px; border-radius: 6px; border: 1px solid #7dd3fc;">
                        <label style="font-size: 11px; color: #334155; font-weight: bold; display: block;">Factory</label>
                        <select id="filter-factory" class="form-control" style="width: 100%; border: none; background: transparent; font-size: 13px; outline: none;" onchange="filterDashboardTable()">
                            <option value="">All</option>
                            ${factories.map(f => `<option value="${f}">${f}</option>`).join('')}
                        </select>
                    </div>
                    <div style="background: #fff; padding: 8px; border-radius: 6px; border: 1px solid #7dd3fc;">
                        <label style="font-size: 11px; color: #334155; font-weight: bold; display: block;">Section</label>
                        <select id="filter-section" class="form-control" style="width: 100%; border: none; background: transparent; font-size: 13px; outline: none;" onchange="filterDashboardTable()">
                            <option value="">All</option>
                            ${sections.map(s => `<option value="${s}">${s}</option>`).join('')}
                        </select>
                    </div>
                    <div style="background: #fff; padding: 8px; border-radius: 6px; border: 1px solid #7dd3fc;">
                        <label style="font-size: 11px; color: #334155; font-weight: bold; display: block;">Style</label>
                        <select id="filter-style" class="form-control" style="width: 100%; border: none; background: transparent; font-size: 13px; outline: none;" onchange="filterDashboardTable()">
                            <option value="">All</option>
                            ${styles.map(st => `<option value="${st}">${st}</option>`).join('')}
                        </select>
                    </div>
                    <div style="background: #fff; padding: 8px; border-radius: 6px; border: 1px solid #7dd3fc;">
                        <label style="font-size: 11px; color: #334155; font-weight: bold; display: block;">Day (D)</label>
                        <select id="filter-rampup" class="form-control" style="width: 100%; border: none; background: transparent; font-size: 13px; outline: none;" onchange="filterDashboardTable()">
                            <option value="">All</option>
                            ${rampupDays.map(r => `<option value="${r}">${r}</option>`).join('')}
                        </select>
                    </div>
                </div>

                <!-- 📊 Graph Visualization Section -->
                <div class="uploaded-files-section" style="margin-top: 25px; background: #fff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <h3 style="margin-bottom: 15px;"><i class="fa-solid fa-chart-bar"></i> ក្រាហ្វបង្ហាញ Cycle Time (CT Average) តាម Seq និង Process</h3>
                    <div style="position: relative; height: 320px; width: 100%;">
                        <canvas id="ctProcessChart"></canvas>
                    </div>
                </div>

                <!-- 🌟 Main Dashboard Table -->
                <div class="uploaded-files-section" style="margin-top: 25px;">
                    <h3><i class="fa-solid fa-list-check"></i> បញ្ជីទិន្នន័យ Time Study និងការវិភាគ (CT & PFM)</h3>
                    <div class="table-responsive" style="margin-top: 15px;">
                        <table id="dashboard-main-table" class="files-table" style="width: 100%; text-align: left;">
                            <thead>
                                <tr>
                                    <th>No</th>
                                    <th>Seq</th>
                                    <th>Start Time</th>
                                    <th>Factory / Title</th>
                                    <th>Section</th>
                                    <th>Style</th>
                                    <th>Day</th>
                                    <th>Sewer</th>
                                    <th>MC</th>
                                    <th>Process (Khmer)</th>
                                    <th>PrimaryKey</th>
                                    <th>SMV</th>
                                    <th>CT (Avg)</th>
                                    <th>Output</th>
                                    <th>PFM (%)</th>
                                    <th>Action</th>
                                </tr>
                            </thead>
                            <tbody id="dashboard-table-body">
                                ${savedStudies.length === 0 ? `
                                    <tr><td colspan="16" style="text-align: center; color: #9ca3af; padding: 20px;">មិនទាន់មានទិន្នន័យ Time Study ត្រូវបាន Save នៅឡើយទេ</td></tr>
                                ` : savedStudies.map((item, index) => {
                                    const pfmVal = parseFloat(item.pfm) || 0;
                                    const outputVal = item.output || 0;
                                    const opSeq = item.seq !== undefined ? item.seq : (item.operatorSeq !== undefined ? item.operatorSeq : 99);
                                    const itemNo = item.no || (index + 1);

                                    let badgeBg = '#dbeafe';
                                    let badgeColor = '#1d4ed8';
                                    if (pfmVal < 84) { badgeBg = '#fee2e2'; badgeColor = '#dc2626'; }
                                    else if (pfmVal >= 85 && pfmVal <= 99) { badgeBg = '#fef3c7'; badgeColor = '#d97706'; }
                                    else if (pfmVal >= 100) { badgeBg = '#e0f2fe'; badgeColor = '#0284c7'; }

                                    return `
                                        <tr data-factory="${item.title || ''}" 
                                            data-section="${item.section || ''}" 
                                            data-style="${item.styleName || ''}" 
                                            data-rampup="${item.rampupDay || ''}">
                                            <td><strong>${itemNo}</strong></td>
                                            <td><span style="background: #f3f4f6; padding: 2px 6px; border-radius: 4px; font-weight: bold;">${opSeq}</span></td>
                                            <td style="font-size: 12px; color: #64748b;">${item.dateTimeStart || item.date}</td>
                                            <td>${item.title || 'N/A'}</td>
                                            <td>${item.section || 'N/A'}</td>
                                            <td>${item.styleName || 'N/A'}</td>
                                            <td>${item.rampupDay || 'N/A'}</td>
                                            <td><strong>${item.sewerName || 'No Sewer'}</strong></td>
                                            <td><span style="background: #f1f5f9; padding: 2px 6px; border-radius: 4px;">${item.mc || '-'}</span></td>
                                            <td><strong>${item.processRegion || 'N/A'}</strong></td>
                                            <td>${item.primaryKey || 'N/A'}</td>
                                            <td>${item.smv || '0.00'}</td>
                                            <td><span style="font-weight: bold; color: #2563eb; background: #eff6ff; padding: 3px 6px; border-radius: 4px;">${item.averageCT || '0.00s'}</span></td>
                                            <td><span style="font-weight: bold; color: #10b981; background: #ecfdf5; padding: 3px 6px; border-radius: 4px;">${outputVal} Pcs</span></td>
                                            <td>
                                                <span style="font-weight: bold; background: ${badgeBg}; color: ${badgeColor}; padding: 4px 8px; border-radius: 6px;">
                                                    ${pfmVal}%
                                                </span>
                                            </td>
                                            <td>
                                                <button onclick="editStudyRecord('${item.id}')" style="background: none; border: none; color: #2563eb; cursor: pointer; margin-right: 5px;" title="Edit Study"><i class="fa-solid fa-pen-to-square"></i></button>
                                                <button onclick="deleteDashboardStudy('${item.id}')" style="background: none; border: none; color: #dc2626; cursor: pointer;" title="Delete"><i class="fa-solid fa-trash"></i></button>
                                            </td>
                                        </tr>
                                    `;
                                }).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        `;

        renderCTChart(savedStudies);
    };

    function renderCTChart(studies) {
        const ctx = document.getElementById('ctProcessChart');
        if (!ctx) return;

        const sewerDataMap = {};
        const sewerSeqMap = {};
        const allProcessesSet = new Set();

        studies.forEach(item => {
            const sewerName = (item.sewerName && item.sewerName !== 'undefined' && item.sewerName.trim() !== '') ? item.sewerName : 'No Sewer';
            const opSeq = parseInt(item.seq) || parseInt(item.operatorSeq) || 99;
            sewerSeqMap[sewerName] = opSeq;

            const processName = item.processRegion || item.processName || 'General Process';
            allProcessesSet.add(processName);

            let durationSec = item.avgSeconds;
            if (durationSec === undefined && item.averageCT) {
                durationSec = parseFloat(item.averageCT.replace('s', '')) || 0;
            }

            if (!sewerDataMap[sewerName]) {
                sewerDataMap[sewerName] = {};
            }
            sewerDataMap[sewerName][processName] = (sewerDataMap[sewerName][processName] || 0) + (durationSec || 0);
        });

        const labels = Object.keys(sewerDataMap).sort((a, b) => {
            return (sewerSeqMap[a] || 99) - (sewerSeqMap[b] || 99);
        });

        const uniqueProcesses = Array.from(allProcessesSet);

        const colorPalette = [
            'rgba(16, 185, 129, 0.85)', 
            'rgba(37, 99, 235, 0.85)', 
            'rgba(245, 158, 11, 0.85)', 
            'rgba(239, 68, 68, 0.85)', 
            'rgba(139, 92, 246, 0.85)', 
            'rgba(236, 72, 153, 0.85)'  
        ];

        const datasets = uniqueProcesses.map((processName, idx) => {
            const data = labels.map(sewerName => {
                const val = sewerDataMap[sewerName][processName];
                return val !== undefined ? Math.round(val) : 0;
            });

            return {
                label: processName,
                data: data,
                backgroundColor: colorPalette[idx % colorPalette.length],
                borderRadius: 4
            };
        });

        if (window.myCTChart instanceof Chart) {
            window.myCTChart.destroy();
        }

        const limitLinesAndValuesPlugin = {
            id: 'limitLinesAndValuesPlugin',
            afterDraw(chart) {
                const { ctx, chartArea: { left, right }, scales: { y } } = chart;
                const limits = [
                    { value: 150, label: 'UCL (150s)', color: '#dc2626' },
                    { value: 90,  label: 'BPT (90s)',  color: '#2563eb' },
                    { value: 30,  label: 'LCL (30s)',  color: '#d97706' }
                ];

                ctx.save();
                limits.forEach(limit => {
                    const yPos = y.getPixelForValue(limit.value);
                    if (yPos >= chart.chartArea.top && yPos <= chart.chartArea.bottom) {
                        ctx.beginPath();
                        ctx.moveTo(left, yPos);
                        ctx.lineTo(right, yPos);
                        ctx.lineWidth = 2;
                        ctx.strokeStyle = limit.color;
                        ctx.setLineDash([6, 6]);
                        ctx.stroke();

                        ctx.fillStyle = limit.color;
                        ctx.font = 'bold 11px sans-serif';
                        ctx.textAlign = 'right';
                        ctx.textBaseline = 'bottom';
                        ctx.fillText(limit.label, right - 5, yPos - 3);
                    }
                });
                ctx.restore();
            },
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta0 = chart.getDatasetMeta(0);
                
                meta0.data.forEach((bar, index) => {
                    let totalSum = 0;
                    chart.data.datasets.forEach(dataset => {
                        totalSum += (dataset.data[index] || 0);
                    });

                    if (totalSum > 0) {
                        ctx.save();
                        ctx.font = 'bold 11px sans-serif';
                        ctx.fillStyle = '#1e293b';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'bottom';

                        const topmostBarMeta = chart.getDatasetMeta(chart.data.datasets.length - 1);
                        if (topmostBarMeta && topmostBarMeta.data[index]) {
                            const position = topmostBarMeta.data[index].tooltipPosition();
                            ctx.fillText(totalSum + 's', position.x, position.y - 6);
                        }
                        ctx.restore();
                    }
                });
            }
        };

        window.myCTChart = new Chart(ctx, {
            type: 'bar',
            data: { labels: labels, datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { top: 30 } },
                scales: {
                    x: { stacked: true, title: { display: true, text: 'Sewer Name (Sorted by Seq)' } },
                    y: { stacked: true, beginAtZero: true, title: { display: true, text: 'Total Seconds (s)' }, suggestedMax: 200 }
                },
                plugins: {
                    legend: { display: true, position: 'top' },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return context.dataset.label + ': ' + context.raw + 's';
                            }
                        }
                    }
                }
            },
            plugins: [limitLinesAndValuesPlugin]
        });
    }

    window.filterDashboardTable = function() {
        const fFactory = document.getElementById('filter-factory').value;
        const fSection = document.getElementById('filter-section').value;
        const fStyle = document.getElementById('filter-style').value;
        const fRampup = document.getElementById('filter-rampup').value;

        const rows = document.querySelectorAll('#dashboard-table-body tr');
        rows.forEach(row => {
            if (!row.getAttribute('data-factory')) return;
            let match = true;
            if (fFactory && row.getAttribute('data-factory') !== fFactory) match = false;
            if (fSection && row.getAttribute('data-section') !== fSection) match = false;
            if (fStyle && row.getAttribute('data-style') !== fStyle) match = false;
            if (fRampup && row.getAttribute('data-rampup') !== fRampup) match = false;
            row.style.display = match ? '' : 'none';
        });
    };

    window.deleteDashboardStudy = function(id) {
        showCustomConfirm('បញ្ជាក់ការលុប', 'តើអ្នកពិតជាចង់លុបទិន្នន័យ Time Study នេះមែនទេ?', () => {
            let savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];
            savedStudies = savedStudies.filter(item => item.id !== id);
            localStorage.setItem('savedTimeStudies', JSON.stringify(savedStudies));
            renderDashboardStats();
            showToast('លុបជោគជ័យ', 'ទិន្នន័យត្រូវបានលុបចេញពីប្រព័ន្ធ', 'fa-solid fa-trash', '#dc2626');
        });
    };

    window.editStudyRecord = function(id) {
        const savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];
        const study = savedStudies.find(item => item.id === id);
        if (!study) return;

        const editIdEl = document.getElementById('edit-study-id');
        if (editIdEl) editIdEl.value = study.id;
        
        const factoryEl = document.getElementById('edit-factory');
        if (factoryEl) factoryEl.value = study.title || study.factory || '';
        
        const sectionEl = document.getElementById('edit-section');
        if (sectionEl) sectionEl.value = study.section || '';
        
        const prepVal = study.prepAssembly || study.prep || study.prep_assembly || '';
        const prepEl = document.getElementById('edit-prep');
        if (prepEl) prepEl.value = prepVal;

        const styleEl = document.getElementById('edit-style');
        if (styleEl) styleEl.value = study.styleName || '';
        
        const rampupEl = document.getElementById('edit-rampup');
        if (rampupEl) rampupEl.value = study.rampupDay || '';
        
        const sewerEl = document.getElementById('edit-sewer');
        if (sewerEl) sewerEl.value = study.sewerName || '';
        
        const seqEl = document.getElementById('edit-seq');
        if (seqEl) seqEl.value = study.operatorSeq !== undefined ? study.operatorSeq : (study.seq || '');
        
        const processEl = document.getElementById('edit-process');
        if (processEl) processEl.value = study.processRegion || '';
        
        const smvEl = document.getElementById('edit-smv');
        if (smvEl) smvEl.value = study.smv || 0;

        const lapsContainer = document.getElementById('edit-laps-container');
        if (lapsContainer) {
            if (study.laps && study.laps.length > 0) {
                lapsContainer.innerHTML = study.laps.map((lap, index) => {
                    const secVal = (lap.durationMs / 1000).toFixed(2);
                    return `
                        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; gap: 8px;">
                            <span style="font-size: 12px; font-weight: bold; color: #475569;">Round ${lap.round || (index + 1)}:</span>
                            <input type="number" step="0.01" class="edit-lap-input" data-index="${index}" value="${secVal}" style="flex: 1; padding: 4px 8px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 13px;">
                            <span style="font-size: 12px; color: #64748b;">s</span>
                        </div>
                    `;
                }).join('');
            } else {
                lapsContainer.innerHTML = '<div style="font-size: 12px; color: #94a3b8; text-align: center;">គ្មានទិន្នន័យ Lap ទេ</div>';
            }
        }

        const editModal = document.getElementById('editStudyModal');
        if (editModal) editModal.style.display = 'flex';
    };

    window.closeEditModal = function() {
        const editModal = document.getElementById('editStudyModal');
        if (editModal) editModal.style.display = 'none';
    };

    window.saveEditedStudy = function() {
        const editIdEl = document.getElementById('edit-study-id');
        if (!editIdEl) return;
        const id = editIdEl.value;
        let savedStudies = JSON.parse(localStorage.getItem('savedTimeStudies')) || [];

        const index = savedStudies.findIndex(item => item.id === id);
        if (index === -1) return;

        savedStudies[index].title = document.getElementById('edit-factory').value;
        savedStudies[index].section = document.getElementById('edit-section').value; 
        
        const newPrepAssembly = document.getElementById('edit-prep').value;
        savedStudies[index].prepAssembly = newPrepAssembly; 

        savedStudies[index].styleName = document.getElementById('edit-style').value;
        savedStudies[index].rampupDay = document.getElementById('edit-rampup').value;
        savedStudies[index].sewerName = document.getElementById('edit-sewer').value;
        savedStudies[index].operatorSeq = parseInt(document.getElementById('edit-seq').value) || 99;
        savedStudies[index].seq = savedStudies[index].operatorSeq;
        savedStudies[index].processRegion = document.getElementById('edit-process').value;
        
        const newSmv = parseFloat(document.getElementById('edit-smv').value) || 0;
        savedStudies[index].smv = newSmv;

        const factoryVal = savedStudies[index].title || '';
        const sectionVal = savedStudies[index].section || '';
        const styleVal = savedStudies[index].styleName || '';
        const rampupVal = savedStudies[index].rampupDay || '';

        savedStudies[index].primaryKey = `${factoryVal}-${sectionVal}-${styleVal}-${rampupVal}-${newPrepAssembly}`;

        const lapInputs = document.querySelectorAll('.edit-lap-input');
        if (lapInputs.length > 0) {
            let newLaps = [];
            let totalMs = 0;

            lapInputs.forEach((input, idx) => {
                const valSec = parseFloat(input.value) || 0;
                const durationMs = Math.round(valSec * 1000);
                totalMs += durationMs;
                newLaps.push({
                    round: idx + 1,
                    durationMs: durationMs
                });
            });

            savedStudies[index].laps = newLaps;
            const avgMs = totalMs / newLaps.length;
            const totalSeconds = avgMs / 1000;
            savedStudies[index].avgSeconds = totalSeconds;

            const secPart = Math.floor(totalSeconds);
            const msPart = Math.floor((avgMs % 1000) / 10);
            savedStudies[index].averageCT = String(secPart).padStart(2, '0') + '.' + String(msPart).padStart(2, '0') + 's';

            if (totalSeconds > 0) {
                savedStudies[index].pfm = ((newSmv / totalSeconds) * 100).toFixed(1);
                savedStudies[index].output = Math.round(3600 / totalSeconds);
            }
        }

        localStorage.setItem('savedTimeStudies', JSON.stringify(savedStudies));

        closeEditModal();
        renderDashboardStats();
        
        if (typeof showToast === 'function') {
            showToast('ជោគជ័យ', 'ទិន្នន័យត្រូវបានផ្លាស់ទីទៅ Prep_Assembly ថ្មីរួចរាល់!', 'fa-solid fa-check', '#10b981');
        }
    };

});

// --- ៧. មុខងារជំនួយ (Global Utility Functions) ---
function showCustomAlert(title, message, iconClass = "fa-solid fa-circle-check", iconColor = "#10b981") {
    const modal = document.getElementById('custom-modal');
    if (!modal) return;
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-message').innerText = message;
    document.getElementById('modal-icon').innerHTML = `<i class="${iconClass}" style="color: ${iconColor};"></i>`;
    const btnCancel = document.getElementById('modal-btn-cancel');
    if (btnCancel) btnCancel.style.display = 'none';
    const btnOk = document.getElementById('modal-btn-ok');
    if (btnOk) {
        btnOk.textContent = 'យល់ព្រម';
        btnOk.className = 'modal-btn modal-btn-primary';
    }
    modal.style.display = 'flex';
    const newBtnOk = btnOk.cloneNode(true);
    btnOk.parentNode.replaceChild(newBtnOk, btnOk);
    newBtnOk.addEventListener('click', () => modal.style.display = 'none');
}

function showCustomConfirm(title, message, onConfirm) {
    let modal = document.getElementById('custom-confirm-modal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'custom-confirm-modal';
        modal.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); display:flex; align-items:center; justify-content:center; z-index:9999;';
        modal.innerHTML = `
            <div style="background:#fff; padding:25px; border-radius:12px; width:350px; text-align:center; box-shadow:0 4px 6px rgba(0,0,0,0.1);">
                <h3 id="confirm-title" style="margin-bottom:10px; color:#1e293b;"></h3>
                <p id="confirm-msg" style="margin-bottom:20px; color:#64748b; font-size:14px;"></p>
                <div style="display:flex; gap:10px; justify-content:center;">
                    <button id="confirm-yes-btn" style="background:#dc2626; color:#fff; border:none; padding:8px 20px; border-radius:6px; cursor:pointer; font-weight:bold;">លុប</button>
                    <button id="confirm-no-btn" style="background:#e2e8f0; color:#334155; border:none; padding:8px 20px; border-radius:6px; cursor:pointer; font-weight:bold;">បោះបង់</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-msg').textContent = message;
    modal.style.display = 'flex';

    document.getElementById('confirm-yes-btn').onclick = function() {
        modal.style.display = 'none';
        if (onConfirm) onConfirm();
    };
    document.getElementById('confirm-no-btn').onclick = function() {
        modal.style.display = 'none';
    };
}

function showToast(title, message, iconClass, iconColor) {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = 'toast success';
    toast.innerHTML = `<i class="${iconClass}" style="color: ${iconColor};"></i><div class="toast-content"><h4>${title}</h4><p>${message}</p></div>`;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}
///5step///
function makeTableResizable(table) {
    if (!table) return;
    table.querySelectorAll('th').forEach((th) => {
        if (th.querySelector('.resizer')) return;
        const resizer = document.createElement('div');
        resizer.classList.add('resizer');
        th.appendChild(resizer);

        let startX = 0, startWidth = 0;
        resizer.addEventListener('mousedown', function (e) {
            startX = e.clientX;
            startWidth = th.offsetWidth;
            th.classList.add('resizing');
            document.addEventListener('mousemove', mouseMoveHandler);
            document.addEventListener('mouseup', mouseUpHandler);
            e.stopPropagation();
            e.preventDefault();
        });

function mouseMoveHandler(e) {
            const dx = e.clientX - startX; // 🌟 ថែមសញ្ញាដក (-) ត្រង់នេះឱ្យបានត្រឹមត្រូវ
            const newWidth = Math.max(40, startWidth + dx);
            th.style.width = newWidth + 'px';
            th.style.minWidth = newWidth + 'px';
            th.style.maxWidth = newWidth + 'px';
        }

        function mouseUpHandler() {
            th.classList.remove('resizing');
            document.removeEventListener('mousemove', mouseMoveHandler);
            document.removeEventListener('mouseup', mouseUpHandler);
        }
    });
}

// 🌟 ធ្វើការ Refresh ទិន្នន័យដោយស្វ័យប្រវត្តិ រាល់ពេលអ្នកប្រើប្រាស់ត្រឡប់មកកាន់ទំព័រនេះវិញ
window.addEventListener('pageshow', function(event) {
    const activeRampupDay = localStorage.getItem('lastSelectedRampupDay');
    const activeTitle = localStorage.getItem('lastSelectedTitle');
    const activeSection = localStorage.getItem('lastSelectedSection');
    const activeStyle = localStorage.getItem('lastSelectedStyle');

    if (activeRampupDay && typeof renderPrepAssemblyList === 'function') {
        renderPrepAssemblyList(activeTitle, activeSection, activeStyle, activeRampupDay);
    }
});