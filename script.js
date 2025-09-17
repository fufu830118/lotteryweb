class LotteryApp {
    constructor() {
        this.currentStep = 1;
        this.excelData = [];
        this.selectedColumn = '';
        this.participants = [];
        this.prizes = [];
        this.winners = {};
        this.isLotteryRunning = false;

        // 移除自動導出和HTML匯出功能

        this.initializeEventListeners();
    }

    initializeEventListeners() {
        // 檔案上傳
        document.getElementById('excelFile').addEventListener('change', (e) => this.handleFileUpload(e));

        // 確認欄位選擇
        document.getElementById('confirmColumn').addEventListener('click', () => this.confirmColumnSelection());

        // 新增獎項
        document.getElementById('addPrize').addEventListener('click', () => this.addPrize());
        document.getElementById('prizeName').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.addPrize();
        });
        document.getElementById('prizeCount').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.addPrize();
        });

        // 開始抽獎
        document.getElementById('startLottery').addEventListener('click', () => this.startLottery());


        // 控制按鈕
        document.getElementById('download-image-btn').addEventListener('click', () => this.downloadScreenshot());
        document.getElementById('exportResults').addEventListener('click', () => this.exportResults());
        document.getElementById('restartLottery').addEventListener('click', () => this.restartApp());
    }

    // Excel 檔案處理
    async handleFileUpload(event) {
        const file = event.target.files[0];
        const errorDiv = document.getElementById('uploadError');
        const fileInfoDiv = document.getElementById('fileInfo');

        // 重設狀態
        errorDiv.classList.add('hidden');
        errorDiv.textContent = '';
        fileInfoDiv.classList.add('hidden');

        if (!file) return;

        try {
            const data = await this.readExcelFile(file);
            this.excelData = data;

            if (this.excelData.length === 0) {
                throw new Error('Excel 檔案中沒有找到有效的資料列。請確認檔案至少包含一個標題列和一筆資料。');
            }

            // 顯示檔案資訊
            document.getElementById('fileName').textContent = file.name;
            document.getElementById('recordCount').textContent = data.length;
            fileInfoDiv.classList.remove('hidden');

            // 生成欄位選項
            this.generateColumnOptions();

            // 進入下一步
            this.nextStep();

        } catch (error) {
            let errorMessage = '❌ 檔案處理失敗：';
            if (error.message.includes('標題列是空的')) {
                errorMessage += '檔案標題列是空的或只包含空格，請確認第一行包含有效的欄位名稱。';
            } else if (error.message.includes('沒有找到有效的資料列')) {
                errorMessage += '檔案中沒有有效資料，請確認檔案包含至少一筆有效資料。';
            } else if (error.message.includes('至少一行標題和一行資料')) {
                errorMessage += '檔案格式不正確，請確認檔案包含標題列和資料列。';
            } else {
                errorMessage += error.message;
            }
            errorDiv.textContent = errorMessage;
            errorDiv.classList.remove('hidden');
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    if (jsonData.length < 2) {
                        reject(new Error('Excel 檔案必須包含至少一行標題和一行資料'));
                        return;
                    }

                    // 轉換為物件格式，清理空白字符
                    const headers = jsonData[0].map(header =>
                        typeof header === 'string' ? header.trim() : header
                    ).filter(header => header !== '');

                    if (headers.length === 0) {
                        reject(new Error('Excel 檔案標題列是空的或只包含空格'));
                        return;
                    }

                    const rows = jsonData.slice(1);
                    const result = rows.map(row => {
                        const obj = {};
                        headers.forEach((header, index) => {
                            const value = row[index];
                            obj[header] = typeof value === 'string' ? value.trim() : (value || '');
                        });
                        return obj;
                    }).filter(row => Object.values(row).some(value =>
                        value !== '' && value !== null && value !== undefined
                    ));

                    if (result.length === 0) {
                        reject(new Error('Excel 檔案中沒有找到有效的資料列（所有列都是空的或只包含空格）'));
                        return;
                    }

                    resolve(result);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('檔案讀取失敗，請確認檔案沒有損壞且為有效的 Excel 格式'));
            reader.readAsArrayBuffer(file);
        });
    }

    generateColumnOptions() {
        if (this.excelData.length === 0) return;

        const headers = Object.keys(this.excelData[0]);
        const container = document.getElementById('columnOptions');
        container.innerHTML = '';

        headers.forEach(header => {
            const option = document.createElement('div');
            option.className = 'column-option';
            option.textContent = header;
            option.addEventListener('click', () => this.selectColumn(header, option));
            container.appendChild(option);
        });
    }

    selectColumn(columnName, element) {
        // 移除其他選項的選中狀態
        document.querySelectorAll('.column-option').forEach(opt => {
            opt.classList.remove('selected');
        });

        // 選中當前選項
        element.classList.add('selected');
        this.selectedColumn = columnName;

        // 顯示預覽資料
        this.showPreviewData();

        // 顯示確認按鈕
        document.getElementById('confirmColumn').classList.remove('hidden');
    }

    showPreviewData() {
        const previewContainer = document.getElementById('dataPreview');
        const previewSection = document.getElementById('previewData');

        // 取前5筆資料預覽
        const previewData = this.excelData.slice(0, 5);

        let html = '<table class="preview-table"><tr>';
        html += `<th>${this.selectedColumn}</th>`;
        html += '</tr>';

        previewData.forEach(row => {
            html += `<tr><td>${row[this.selectedColumn] || ''}</td></tr>`;
        });

        html += '</table>';
        previewContainer.innerHTML = html;
        previewSection.classList.remove('hidden');
    }

    confirmColumnSelection() {
        // 建立參與者列表，更嚴格的空值檢查
        this.participants = this.excelData
            .map(row => row[this.selectedColumn])
            .filter(name => {
                // 檢查空值、空字符串、純空格
                if (name === null || name === undefined) return false;
                const trimmed = name.toString().trim();
                return trimmed !== '';
            })
            .map(name => name.toString().trim());

        if (this.participants.length === 0) {
            const errorDiv = document.getElementById('uploadError');
            errorDiv.textContent = `❌ 選擇的欄位「${this.selectedColumn}」中沒有有效的資料！請選擇其他欄位或檢查 Excel 檔案。`;
            errorDiv.classList.remove('hidden');
            return;
        }

        // 清除錯誤訊息
        const errorDiv = document.getElementById('uploadError');
        errorDiv.classList.add('hidden');

        this.nextStep();
    }

    // 獎項管理
    addPrize() {
        const nameInput = document.getElementById('prizeName');
        const countInput = document.getElementById('prizeCount');

        const name = nameInput.value.trim();
        const count = parseInt(countInput.value);

        if (!name) {
            alert('❌ 請輸入獎項名稱！');
            nameInput.focus();
            return;
        }

        if (!count || count < 1) {
            alert('❌ 請輸入有效的獎項數量！（必須大於 0）');
            countInput.focus();
            return;
        }

        if (count > this.participants.length) {
            alert(`❌ 獎項數量（${count}個）不能超過參與者總數（${this.participants.length}人）！`);
            countInput.focus();
            return;
        }

        // 檢查是否已存在相同名稱的獎項
        if (this.prizes.some(prize => prize.name === name)) {
            alert(`❌ 獎項名稱「${name}」已存在！請使用不同的名稱。`);
            nameInput.focus();
            return;
        }

        this.prizes.push({ name, count });

        // 清空輸入框
        nameInput.value = '';
        countInput.value = '';
        nameInput.focus();

        this.updatePrizeList();
    }

    updatePrizeList() {
        const container = document.getElementById('prizeList');

        if (this.prizes.length === 0) {
            container.innerHTML = '<p style="color: #666; text-align: center;">尚未新增任何獎項</p>';
            document.getElementById('startLottery').classList.add('hidden');
            return;
        }

        let html = '';
        this.prizes.forEach((prize, index) => {
            html += `
                <div class="prize-item">
                    <div class="prize-info">
                        <div class="prize-name">${prize.name}</div>
                        <div class="prize-count">數量：${prize.count} 個</div>
                    </div>
                    <button class="delete-prize" onclick="app.deletePrize(${index})">刪除</button>
                </div>
            `;
        });

        container.innerHTML = html;

        // 檢查總獎項數量
        const totalPrizes = this.prizes.reduce((sum, prize) => sum + prize.count, 0);
        if (totalPrizes <= this.participants.length) {
            document.getElementById('startLottery').classList.remove('hidden');
        } else {
            document.getElementById('startLottery').classList.add('hidden');
            alert(`❌ 總獎項數量（${totalPrizes}個）不能超過參與者總數（${this.participants.length}人）！`);
        }
    }

    deletePrize(index) {
        this.prizes.splice(index, 1);
        this.updatePrizeList();
    }

    async startLottery() {
        this.nextStep();
        this.isLotteryRunning = true;

        document.getElementById('currentPrizeName').textContent = '🎲 正在計算抽獎結果...';
        await this.delay(500);

        const allResults = this.drawAllPrizesAtOnce();
        if (!allResults) {
            this.isLotteryRunning = false;
            return;
        }
        this.winners = allResults;

        // 依序將結果帶動畫顯示在下方列表
        for (let i = 0; i < this.prizes.length; i++) {
            const prize = this.prizes[i];
            const winners = this.winners[prize.name];

            document.getElementById('currentPrizeName').textContent = `🎁 正在揭曉：${prize.name}`;
            
            await this.displayPrizeResults(prize.name, winners);
            await this.delay(1000); // 每個獎項之間的停頓
        }

        this.isLotteryRunning = false;
        document.getElementById('currentPrizeName').textContent = '🎉 所有獎項抽獎完成！';
        document.getElementById('lotteryControls').classList.remove('hidden');

        // 保存結果到 localStorage 供獨立頁面使用
        localStorage.setItem('lotteryResults', JSON.stringify(this.winners));

    }

    drawAllPrizesAtOnce() {
        const totalPrizes = this.prizes.reduce((sum, prize) => sum + prize.count, 0);

        if (totalPrizes > this.participants.length) {
            alert(`錯誤：總獎項數量 (${totalPrizes}) 超過參與者人數 (${this.participants.length})！`);
            return null;
        }

        // 洗牌算法：確保完全隨機且不重複
        const shuffledParticipants = this.shuffleArray([...this.participants]);

        const results = {};
        let currentIndex = 0;

        // 依序分配每個獎項的得獎者
        this.prizes.forEach(prize => {
            results[prize.name] = shuffledParticipants.slice(currentIndex, currentIndex + prize.count);
            currentIndex += prize.count;
        });

        return results;
    }

    shuffleArray(array) {
        const shuffled = [...array];
        for (let i = shuffled.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
        }
        return shuffled;
    }

    async displayPrizeResults(prizeName, winners) {
        const winnersList = document.getElementById('winnersList');

        // 建立這個獎項的容器
        const prizeGroup = document.createElement('div');
        prizeGroup.className = 'prize-group animated';

        const prizeTitle = document.createElement('h3');
        prizeTitle.textContent = `${prizeName} (${winners.length}名)`;
        prizeGroup.appendChild(prizeTitle);

        const winnersContainer = document.createElement('div');
        winnersContainer.className = 'winners';
        prizeGroup.appendChild(winnersContainer);

        winnersList.appendChild(prizeGroup);

        // 逐一顯示得獎者
        for (let i = 0; i < winners.length; i++) {
            const winnerName = winners[i];
            const winnerDiv = document.createElement('div');
            winnerDiv.className = 'winner';
            winnerDiv.textContent = winnerName;

            winnersContainer.appendChild(winnerDiv);

            // 讓整個頁面滾動到新得獎者位置
            winnerDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });

            await this.delay(200);
        }
    }

    async exportResults() {
        if (Object.keys(this.winners).length === 0) {
            alert('沒有抽獎結果可以匯出！');
            return;
        }

        // 詢問活動名稱
        const activityName = prompt('請輸入活動名稱（將用於檔案名稱）:', localStorage.getItem('activityName') || 'Wiwynn抽獎活動');
        if (!activityName) {
            return; // 使用者取消
        }

        // 保存活動名稱
        localStorage.setItem('activityName', activityName);

        // 1. 匯出 Excel
        try {
            const exportData = [];
            Object.entries(this.winners).forEach(([prizeName, winners]) => {
                winners.forEach((winner, index) => {
                    exportData.push({
                        '活動名稱': activityName,
                        '獎項': prizeName,
                        '序號': index + 1,
                        '得獎者': winner
                    });
                });
            });

            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '抽獎結果');
            const sanitizedActivityName = activityName.replace(/[\\/:*?"<>|]/g, '_');
            const excelFileName = `${sanitizedActivityName}_抽獎結果_${new Date().toISOString().slice(0, 10)}.xlsx`;
            XLSX.writeFile(wb, excelFileName);
        } catch (error) {
            console.error('匯出 Excel 失敗:', error);
            alert('匯出 Excel 失敗，請檢查主控台錯誤訊息。');
        }
    }

    async downloadScreenshot() {
        if (Object.keys(this.winners).length === 0) {
            alert('沒有抽獎結果可以截圖！');
            return;
        }

        // 詢問活動名稱
        const activityName = prompt('請輸入活動名稱:', localStorage.getItem('activityName') || 'Wiwynn抽獎活動');
        if (!activityName) {
            return; // 使用者取消
        }

        // 保存活動名稱和結果到 localStorage 供 results.html 讀取
        localStorage.setItem('activityName', activityName);
        localStorage.setItem('lotteryResults', JSON.stringify(this.winners));

        // 直接開啟 results.html 頁面，附帶活動名稱參數
        const encodedActivity = encodeURIComponent(activityName);
        window.open(`results.html?activity=${encodedActivity}`, '_blank');
    }

    async waitForAnimationsToComplete() {
        return new Promise((resolve) => {
            // 檢查所有動畫元素
            const animatedElements = document.querySelectorAll('.winner, .prize-group.animated');

            if (animatedElements.length === 0) {
                resolve();
                return;
            }

            let completedAnimations = 0;
            const totalAnimations = animatedElements.length;

            const checkCompletion = () => {
                completedAnimations++;
                if (completedAnimations >= totalAnimations) {
                    // 額外等待一點時間確保所有動畫完全結束
                    setTimeout(resolve, 500);
                }
            };

            animatedElements.forEach(element => {
                // 監聽動畫結束事件
                element.addEventListener('animationend', checkCompletion, { once: true });
                element.addEventListener('transitionend', checkCompletion, { once: true });
            });

            // 設置超時以防動畫事件不觸發
            setTimeout(() => {
                resolve();
            }, 5000);
        });
    }



    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    nextStep() {
        // 隱藏當前步驟
        document.getElementById(`step${this.currentStep}`).classList.add('hidden');
        document.querySelector(`.step[data-step="${this.currentStep}"]`).classList.remove('active');

        // 顯示下一步驟
        this.currentStep++;
        document.getElementById(`step${this.currentStep}`).classList.remove('hidden');
        document.querySelector(`.step[data-step="${this.currentStep}"]`).classList.add('active');
    }

    restartApp() {
        if (confirm('⚠️ 確定要重新開始嗎？\n\n這將清除所有資料，包括：\n• 已匯入的參與者名單\n• 設定的獎項\n• 抽獎結果\n\n此操作無法復原！')) {
            location.reload();
        }
    }
}

// 初始化應用程式
const app = new LotteryApp();

// 防止頁面意外關閉時遺失資料
window.addEventListener('beforeunload', (e) => {
    if (app.participants.length > 0 || app.prizes.length > 0) {
        e.preventDefault();
        e.returnValue = '';
    }
});

// 錯誤處理
window.addEventListener('error', (e) => {
    console.error('發生錯誤：', e.error);
    alert('系統發生錯誤，請重新整理頁面後再試。');
});

// 拖拽上傳支援
const uploadBox = document.querySelector('.upload-box');
if (uploadBox) {
    uploadBox.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadBox.style.borderColor = '#667eea';
        uploadBox.style.background = 'rgba(102, 126, 234, 0.1)';
    });

    uploadBox.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadBox.style.borderColor = '#ddd';
        uploadBox.style.background = 'transparent';
    });

    uploadBox.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadBox.style.borderColor = '#ddd';
        uploadBox.style.background = 'transparent';

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const fileInput = document.getElementById('excelFile');
            fileInput.files = files;
            fileInput.dispatchEvent(new Event('change'));
        }
    });
}