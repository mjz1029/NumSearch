import os
import sys
import threading
import webbrowser
import time
from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import socket

# 前端HTML内容（内嵌，无需外部文件）
HTML_CONTENT = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>手机号码归属地批量查询工具</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://cdn.jsdelivr.net/npm/font-awesome@4.7.0/css/font-awesome.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        primary: '#3B82F6',
                        secondary: '#10B981',
                        accent: '#6366F1',
                        neutral: '#1F2937',
                        light: '#F3F4F6',
                    },
                    fontFamily: {
                        sans: ['Inter', 'system-ui', 'sans-serif'],
                    },
                }
            }
        }
    </script>
    <style type="text/tailwindcss">
        @layer utilities {
            .content-auto {
                content-visibility: auto;
            }
            .shadow-soft {
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
            }
            .transition-custom {
                transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            }
        }
    </style>
</head>
<body class="bg-gray-50 min-h-screen font-sans text-neutral">
    <div class="container mx-auto px-4 py-8 max-w-5xl">
        <header class="text-center mb-10">
            <h1 class="text-[clamp(1.8rem,5vw,2.8rem)] font-bold text-neutral mb-2 bg-gradient-to-r from-primary to-accent text-transparent bg-clip-text">
                手机号码归属地批量查询工具
            </h1>
            <p class="text-gray-600 max-w-2xl mx-auto">
                灵活版 | 支持自定义电话列，自动添加表头
            </p>
        </header>

        <div id="server-status" class="bg-green-50 border-l-4 border-green-400 p-4 mb-6">
            <div class="flex">
                <div class="flex-shrink-0">
                    <i class="fa fa-server text-green-500"></i>
                </div>
                <div class="ml-3">
                    <p class="text-sm text-green-700" id="server-status-text">
                        <i class="fa fa-check-circle text-green-500 mr-1"></i>服务已启动，可正常使用
                    </p>
                </div>
            </div>
        </div>

        <main class="bg-white rounded-xl shadow-soft p-6 md:p-8 mb-8">
            <div id="upload-section" class="mb-8">
                <div id="drop-area" class="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center transition-custom hover:border-primary cursor-pointer">
                    <i class="fa fa-cloud-upload text-5xl text-gray-400 mb-4 transition-custom hover:text-primary"></i>
                    <h2 class="text-xl font-semibold mb-2">拖放Excel文件到此处</h2>
                    <p class="text-gray-500 mb-4">或点击选择文件</p>
                    <label class="inline-block bg-primary hover:bg-primary/90 text-white font-medium py-2 px-6 rounded-lg transition-custom cursor-pointer">
                        <i class="fa fa-file-excel-o mr-2"></i>选择Excel文件
                        <input type="file" id="file-input" accept=".xlsx, .xls" class="hidden">
                    </label>
                    <p class="text-sm text-gray-400 mt-4">支持格式: .xlsx, .xls</p>
                </div>

                <div id="file-info" class="hidden mt-4 p-4 bg-light rounded-lg">
                    <div class="flex items-center justify-between">
                        <div class="flex items-center">
                            <i class="fa fa-file-excel-o text-secondary text-xl mr-3"></i>
                            <div>
                                <p id="file-name" class="font-medium"></p>
                                <p id="file-size" class="text-sm text-gray-500"></p>
                            </div>
                        </div>
                        <button id="remove-file" class="text-gray-400 hover:text-red-500 transition-custom">
                            <i class="fa fa-times-circle text-xl"></i>
                        </button>
                    </div>
                </div>
            </div>

            <div id="settings-section" class="mb-8 hidden">
                <div class="bg-light rounded-lg p-5">
                    <h3 class="text-lg font-semibold mb-4 flex items-center">
                        <i class="fa fa-cog text-primary mr-2"></i>查询设置
                    </h3>

                    <div class="space-y-4">
                        <div>
                            <label for="phone-column" class="block text-sm font-medium text-gray-700 mb-1">手机号码所在列</label>
                            <select id="phone-column" class="w-full p-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary focus:border-primary">
                                <option value="">自动识别（表头包含"联系电话"）</option>
                                <!-- 列选项会在文件上传后动态生成 -->
                            </select>
                            <p class="text-xs text-gray-500 mt-1">选择Excel中存放手机号码的列</p>
                        </div>

                        <div>
                            <label for="concurrent" class="block text-sm font-medium text-gray-700 mb-1">并发查询数量</label>
                            <div class="flex items-center">
                                <input type="range" id="concurrent" min="5" max="30" value="15" 
                                    class="w-full h-2 bg-gray-200 rounded-lg appearance-none cursor-pointer accent-primary">
                                <span id="concurrent-value" class="ml-3 min-w-[3rem] text-center font-medium">15</span>
                            </div>
                            <p class="text-xs text-gray-500 mt-1">推荐设置：10-20</p>
                        </div>
                    </div>
                </div>
            </div>

            <div id="action-buttons" class="flex justify-center mb-8 hidden">
                <button id="start-process" class="bg-secondary hover:bg-secondary/90 text-white font-medium py-3 px-8 rounded-lg transition-custom flex items-center shadow-lg hover:shadow-xl">
                    <i class="fa fa-play-circle mr-2"></i>开始批量查询
                </button>
            </div>

            <div id="progress-section" class="hidden mb-8">
                <div class="bg-light rounded-lg p-5">
                    <h3 class="text-lg font-semibold mb-4 flex items-center">
                        <i class="fa fa-spinner fa-spin text-primary mr-2"></i>处理进度
                    </h3>

                    <div class="mb-2 flex justify-between text-sm">
                        <span>总进度</span>
                        <span id="overall-progress-text">0/0</span>
                    </div>
                    <div class="w-full bg-gray-200 rounded-full h-2.5 mb-6">
                        <div id="overall-progress-bar" class="bg-primary h-2.5 rounded-full transition-all duration-300" style="width: 0%"></div>
                    </div>

                    <div class="grid md:grid-cols-3 gap-4 mb-4">
                        <div class="bg-white p-3 rounded-lg shadow-sm">
                            <p class="text-sm text-gray-500">已完成</p>
                            <p id="completed-count" class="text-xl font-bold text-secondary">0</p>
                        </div>
                        <div class="bg-white p-3 rounded-lg shadow-sm">
                            <p class="text-sm text-gray-500">处理中</p>
                            <p id="processing-count" class="text-xl font-bold text-primary">0</p>
                        </div>
                        <div class="bg-white p-3 rounded-lg shadow-sm">
                            <p class="text-sm text-gray-500">错误</p>
                            <p id="error-count" class="text-xl font-bold text-red-500">0</p>
                        </div>
                    </div>

                    <div class="text-sm text-gray-500 mb-2">最近处理:</div>
                    <div id="recent-log" class="text-sm h-20 overflow-y-auto bg-white p-3 rounded-lg text-gray-700"></div>
                </div>
            </div>

            <div id="result-section" class="hidden mb-8">
                <div class="bg-light rounded-lg p-5 text-center">
                    <div class="inline-flex items-center justify-center w-16 h-16 rounded-full bg-green-100 text-secondary mb-4">
                        <i class="fa fa-check text-2xl"></i>
                    </div>
                    <h3 class="text-xl font-semibold mb-2">处理完成!</h3>
                    <p class="text-gray-600 mb-6">已成功查询所有手机号码的归属地信息</p>

                    <button id="download-result" class="bg-accent hover:bg-accent/90 text-white font-medium py-3 px-8 rounded-lg transition-custom flex items-center mx-auto shadow-lg hover:shadow-xl">
                        <i class="fa fa-download mr-2"></i>下载结果文件
                    </button>

                    <button id="process-another" class="mt-4 text-primary hover:text-primary/80 font-medium transition-custom">
                        <i class="fa fa-refresh mr-1"></i>处理另一个文件
                    </button>
                </div>
            </div>
        </main>

        <section class="bg-white rounded-xl shadow-soft p-6 mb-8">
            <h3 class="text-lg font-semibold mb-4 flex items-center">
                <i class="fa fa-info-circle text-primary mr-2"></i>使用说明
            </h3>
            <div class="text-gray-600 space-y-2 text-sm">
                <p>1. 支持任意Excel格式，自动识别或手动选择手机号码所在列</p>
                <p>2. 联系电话列后两列若为空，会自动添加"归属地"和"运营商"表头</p>
                <p>3. 22位号码会自动拆分为两行，保留原始行其他数据</p>
                <p>4. 所有查询结果会强制写入联系电话列的后两列，确保数据位置正确</p>
            </div>
        </section>

        <footer class="text-center text-gray-500 text-sm py-4">
            <p>手机号码归属地查询工具 &copy; 2023</p>
        </footer>
    </div>

    <script>
        let selectedFile = null;
        let workbook = null;
        let processedWorkbook = null;
        let originalData = []; // 保存原始数据用于处理
        let concurrentCount = 15;
        let phoneColumn = ''; // 手机号码所在列，空表示自动识别
        let totalNumbers = 0; // 总手机号码数量（包括拆分后的）
        let processedNumbers = 0; // 已处理的手机号码数量
        let locationColumn; // 归属地列索引
        let operatorColumn; // 运营商列索引

        // 代理地址会由后端自动注入
        const localProxyUrl = '__PROXY_URL__';

        // DOM元素
        const dropArea = document.getElementById('drop-area');
        const fileInput = document.getElementById('file-input');
        const fileInfo = document.getElementById('file-info');
        const fileName = document.getElementById('file-name');
        const fileSize = document.getElementById('file-size');
        const removeFile = document.getElementById('remove-file');
        const settingsSection = document.getElementById('settings-section');
        const actionButtons = document.getElementById('action-buttons');
        const concurrentSlider = document.getElementById('concurrent');
        const concurrentValue = document.getElementById('concurrent-value');
        const phoneColumnSelect = document.getElementById('phone-column');
        const startProcess = document.getElementById('start-process');
        const progressSection = document.getElementById('progress-section');
        const overallProgressBar = document.getElementById('overall-progress-bar');
        const overallProgressText = document.getElementById('overall-progress-text');
        const completedCount = document.getElementById('completed-count');
        const processingCount = document.getElementById('processing-count');
        const errorCount = document.getElementById('error-count');
        const recentLog = document.getElementById('recent-log');
        const resultSection = document.getElementById('result-section');
        const downloadResult = document.getElementById('download-result');
        const processAnother = document.getElementById('process-another');

        function initEventListeners() {
            dropArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                dropArea.classList.add('border-primary', 'bg-blue-50');
            });

            dropArea.addEventListener('dragleave', () => {
                dropArea.classList.remove('border-primary', 'bg-blue-50');
            });

            dropArea.addEventListener('drop', (e) => {
                e.preventDefault();
                dropArea.classList.remove('border-primary', 'bg-blue-50');
                if (e.dataTransfer.files.length) {
                    handleFile(e.dataTransfer.files[0]);
                }
            });

            dropArea.addEventListener('click', () => {
                if (!selectedFile) fileInput.click();
            });

            fileInput.addEventListener('change', () => {
                if (fileInput.files.length) handleFile(fileInput.files[0]);
            });

            removeFile.addEventListener('click', resetFileSelection);
            concurrentSlider.addEventListener('input', (e) => {
                concurrentCount = parseInt(e.target.value);
                concurrentValue.textContent = concurrentCount;
            });

            startProcess.addEventListener('click', startProcessing);
            downloadResult.addEventListener('click', downloadResultFile);
            processAnother.addEventListener('click', resetAll);
        }

        function handleFile(file) {
            if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
                alert('请选择Excel文件（.xlsx 或 .xls格式）');
                return;
            }

            selectedFile = file;
            fileName.textContent = file.name;
            fileSize.textContent = formatFileSize(file.size);
            fileInfo.classList.remove('hidden');

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    workbook = XLSX.read(data, { type: 'array' });

                    // 获取第一个工作表
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    originalData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    // 生成列选项
                    generateColumnOptions();

                    // 显示设置区域和操作按钮
                    settingsSection.classList.remove('hidden');
                    actionButtons.classList.remove('hidden');

                } catch (error) {
                    alert('无法解析Excel文件，请检查格式');
                    console.error('Excel解析错误:', error);
                    resetFileSelection();
                }
            };
            reader.readAsArrayBuffer(file);
        }

        // 生成列选项（A, B, C, ...）
        function generateColumnOptions() {
            // 清空现有选项（保留第一个"自动识别"选项）
            while (phoneColumnSelect.options.length > 1) {
                phoneColumnSelect.remove(1);
            }

            // 计算最大列数
            const maxColumns = originalData.length > 0 ? Math.max(...originalData.map(row => row.length)) : 10;

            // 添加列选项
            for (let i = 0; i < maxColumns; i++) {
                const option = document.createElement('option');
                const columnLetter = columnIndexToLetter(i);
                // 尝试获取表头文本
                let headerText = originalData.length > 0 && originalData[0][i] ? 
                                 originalData[0][i].toString().substring(0, 20) : '';
                option.value = i;
                option.textContent = `${columnLetter}列${headerText ? ` (${headerText})` : ''}`;
                phoneColumnSelect.appendChild(option);
            }

            // 尝试自动识别"联系电话"列
            if (originalData.length > 0) {
                const headerRow = originalData[0];
                for (let i = 0; i < headerRow.length; i++) {
                    const headerText = headerRow[i] ? headerRow[i].toString().toLowerCase() : '';
                    if (headerText.includes('电话') || headerText.includes('手机') || headerText.includes('contact')) {
                        phoneColumnSelect.value = i;
                        break;
                    }
                }
            }
        }

        function resetFileSelection() {
            selectedFile = null;
            workbook = null;
            originalData = [];
            fileInput.value = '';
            fileInfo.classList.add('hidden');
            settingsSection.classList.add('hidden');
            actionButtons.classList.add('hidden');
        }

        function resetAll() {
            resetFileSelection();
            progressSection.classList.add('hidden');
            resultSection.classList.add('hidden');
        }

        function startProcessing() {
            if (!workbook || !selectedFile || originalData.length === 0) return;

            // 获取用户选择的电话列
            const selectedPhoneColumn = phoneColumnSelect.value;
            if (selectedPhoneColumn === '') {
                // 自动识别"联系电话"列
                let found = false;
                if (originalData.length > 0) {
                    const headerRow = originalData[0];
                    for (let i = 0; i < headerRow.length; i++) {
                        const headerText = headerRow[i] ? headerRow[i].toString().toLowerCase() : '';
                        if (headerText.includes('联系电话') || headerText.includes('手机号码')) {
                            phoneColumn = i;
                            found = true;
                            break;
                        }
                    }
                }

                if (!found) {
                    alert('未找到包含"联系电话"或"手机号码"的列，请手动选择');
                    return;
                }
            } else {
                phoneColumn = parseInt(selectedPhoneColumn);
            }

            // 计算归属地和运营商列索引（联系电话列的后两列）
            locationColumn = phoneColumn + 1;
            operatorColumn = phoneColumn + 2;

            // 确保表头行存在（至少有一行）
            if (originalData.length === 0) {
                originalData.push([]); // 添加空表头行
            }

            // 自动添加表头（如果为空）
            addMissingHeaders();

            // 显示进度区域
            actionButtons.classList.add('hidden');
            progressSection.classList.remove('hidden');
            resultSection.classList.add('hidden');

            // 重置进度数据
            overallProgressBar.style.width = '0%';
            totalNumbers = countTotalNumbers();
            processedNumbers = 0;
            overallProgressText.textContent = `0/${totalNumbers}`;
            completedCount.textContent = '0';
            processingCount.textContent = '0';
            errorCount.textContent = '0';
            recentLog.innerHTML = '';

            // 创建结果工作簿的深拷贝
            processedWorkbook = JSON.parse(JSON.stringify(workbook));

            // 更新工作簿中的表头
            updateWorkbookHeaders();

            // 开始处理数据
            processWorkbook();
        }

        // 添加缺失的"归属地"和"运营商"表头
        function addMissingHeaders() {
            const headerRow = originalData[0]; // 表头行

            // 确保表头行有足够的列
            while (headerRow.length <= operatorColumn) {
                headerRow.push('');
            }

            // 如果归属地列头为空，添加默认表头
            if (!headerRow[locationColumn] || headerRow[locationColumn].toString().trim() === '') {
                headerRow[locationColumn] = '归属地';
            }

            // 如果运营商列头为空，添加默认表头
            if (!headerRow[operatorColumn] || headerRow[operatorColumn].toString().trim() === '') {
                headerRow[operatorColumn] = '运营商';
            }
        }

        // 更新工作簿中的表头
        function updateWorkbookHeaders() {
            const firstSheetName = workbook.SheetNames[0];
            const headerRow = originalData[0];

            // 写入归属地表头
            const locationColLetter = columnIndexToLetter(locationColumn);
            processedWorkbook.Sheets[firstSheetName][`${locationColLetter}1`] = {
                v: headerRow[locationColumn]
            };

            // 写入运营商表头
            const operatorColLetter = columnIndexToLetter(operatorColumn);
            processedWorkbook.Sheets[firstSheetName][`${operatorColLetter}1`] = {
                v: headerRow[operatorColumn]
            };

            // 更新工作表范围（确保能包含新添加的列）
            const sheet = processedWorkbook.Sheets[firstSheetName];
            const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
            if (range.e.c < operatorColumn) {
                range.e.c = operatorColumn;
                sheet['!ref'] = XLSX.utils.encode_range(range);
            }
        }

        // 计算总手机号码数量（包括拆分后的）
        function countTotalNumbers() {
            let count = 0;
            // 从第1行开始（跳过表头）
            for (let i = 1; i < originalData.length; i++) {
                const row = originalData[i];
                const phoneValue = row[phoneColumn] || '';
                const cleanedPhone = cleanPhoneNumber(phoneValue.toString());

                if (cleanedPhone.length === 11) {
                    count++;
                } else if (cleanedPhone.length === 22) {
                    // 22位视为两个号码
                    count += 2;
                } else if (cleanedPhone.length > 0) {
                    // 其他长度视为一个号码（可能无效）
                    count++;
                }
            }
            return count;
        }

        function processWorkbook() {
            const firstSheetName = workbook.SheetNames[0];

            // 准备需要处理的所有号码（包括拆分后的）
            const processingQueue = [];

            // 从第1行开始（跳过表头）
            for (let rowIndex = 1; rowIndex < originalData.length; rowIndex++) {
                const originalRowIndex = rowIndex;
                const row = originalData[rowIndex];
                const phoneValue = row[phoneColumn] || '';
                const cleanedPhone = cleanPhoneNumber(phoneValue.toString());

                if (cleanedPhone.length === 11) {
                    // 11位号码：正常处理
                    processingQueue.push({
                        originalRowIndex,
                        phoneNumber: cleanedPhone,
                        isFirstOfPair: true, // 是否为一对号码中的第一个
                        pairIndex: 0 // 0表示单个号码，1和2表示一对中的两个
                    });
                } else if (cleanedPhone.length === 22) {
                    // 22位号码：拆分为两个11位号码
                    const firstNumber = cleanedPhone.substring(0, 11);
                    const secondNumber = cleanedPhone.substring(11, 22);

                    processingQueue.push({
                        originalRowIndex,
                        phoneNumber: firstNumber,
                        isFirstOfPair: true,
                        pairIndex: 1
                    });

                    processingQueue.push({
                        originalRowIndex,
                        phoneNumber: secondNumber,
                        isFirstOfPair: false,
                        pairIndex: 2
                    });
                } else if (cleanedPhone.length > 0) {
                    // 其他长度的号码：尝试处理
                    processingQueue.push({
                        originalRowIndex,
                        phoneNumber: cleanedPhone,
                        isFirstOfPair: true,
                        pairIndex: 0
                    });
                }
            }

            // 并发处理队列
            processQueueInBatches(processingQueue, firstSheetName);
        }

        // 批量处理队列
        function processQueueInBatches(queue, sheetName) {
            const batchSize = concurrentCount;
            let currentIndex = 0;

            // 处理单个批次
            function processBatch() {
                if (currentIndex >= queue.length) {
                    // 所有批次处理完成
                    progressSection.classList.add('hidden');
                    resultSection.classList.remove('hidden');
                    return;
                }

                const batch = queue.slice(currentIndex, currentIndex + batchSize);
                currentIndex += batchSize;

                const promises = batch.map(item => {
                    return new Promise(async (resolve) => {
                        processingCount.textContent = parseInt(processingCount.textContent) + 1;

                        try {
                            let location = '';
                            let operator = '';

                            if (item.phoneNumber.length === 11 && /^\d{11}$/.test(item.phoneNumber)) {
                                // 有效的11位数字手机号
                                const result = await getPhoneInfo(item.phoneNumber);
                                location = result.location;
                                operator = result.operator;
                                addLog(`成功: ${item.phoneNumber} → ${location}, ${operator}`);
                            } else {
                                // 无效手机号
                                location = '无效手机号';
                                operator = '无效手机号';
                                addLog(`失败: ${item.phoneNumber} → 无效手机号`);
                                errorCount.textContent = parseInt(errorCount.textContent) + 1;
                            }

                            // 计算需要写入的行索引
                            // 如果是一对号码中的第二个，需要插入新行
                            let rowToWrite = item.originalRowIndex + 1; // +1 因为Excel行索引从1开始

                            if (!item.isFirstOfPair) {
                                // 插入新行
                                insertRowInWorkbook(sheetName, rowToWrite);
                                // 复制原始行数据到新行
                                copyRowData(item.originalRowIndex, rowToWrite - 1);
                            }

                            // 归属地和运营商列
                            const locationColLetter = columnIndexToLetter(locationColumn);
                            const operatorColLetter = columnIndexToLetter(operatorColumn);

                            // 写入归属地（强制写入，覆盖原有数据）
                            processedWorkbook.Sheets[sheetName][`${locationColLetter}${rowToWrite}`] = {
                                v: location
                            };

                            // 写入运营商（强制写入，覆盖原有数据）
                            processedWorkbook.Sheets[sheetName][`${operatorColLetter}${rowToWrite}`] = {
                                v: operator
                            };

                            // 如果是拆分后的号码，更新电话号码列
                            if (item.pairIndex === 1) {
                                const phoneColLetter = columnIndexToLetter(phoneColumn);
                                processedWorkbook.Sheets[sheetName][`${phoneColLetter}${rowToWrite}`] = {
                                    v: item.phoneNumber
                                };
                            } else if (item.pairIndex === 2) {
                                const phoneColLetter = columnIndexToLetter(phoneColumn);
                                processedWorkbook.Sheets[sheetName][`${phoneColLetter}${rowToWrite}`] = {
                                    v: item.phoneNumber
                                };
                            }

                        } catch (error) {
                            const errorMsg = error.message || '查询失败';
                            addLog(`失败: ${item.phoneNumber} → ${errorMsg}`);
                            errorCount.textContent = parseInt(errorCount.textContent) + 1;

                            // 即使出错也写入错误信息
                            const locationColLetter = columnIndexToLetter(locationColumn);
                            const operatorColLetter = columnIndexToLetter(operatorColumn);
                            let rowToWrite = item.originalRowIndex + 1;

                            if (!item.isFirstOfPair) {
                                insertRowInWorkbook(sheetName, rowToWrite);
                                copyRowData(item.originalRowIndex, rowToWrite - 1);
                            }

                            processedWorkbook.Sheets[sheetName][`${locationColLetter}${rowToWrite}`] = {
                                v: errorMsg
                            };

                            processedWorkbook.Sheets[sheetName][`${operatorColLetter}${rowToWrite}`] = {
                                v: errorMsg
                            };
                        } finally {
                            processedNumbers++;
                            completedCount.textContent = processedNumbers;
                            processingCount.textContent = parseInt(processingCount.textContent) - 1;

                            // 更新进度条
                            const progress = Math.round((processedNumbers / totalNumbers) * 100);
                            overallProgressBar.style.width = `${progress}%`;
                            overallProgressText.textContent = `${processedNumbers}/${totalNumbers}`;

                            resolve();
                        }
                    });
                });

                // 等待当前批次完成后再处理下一批
                Promise.all(promises).then(processBatch);
            }

            // 开始处理第一批
            processBatch();
        }

        // 在工作簿中插入新行
        function insertRowInWorkbook(sheetName, rowIndex) {
            const sheet = processedWorkbook.Sheets[sheetName];

            // 确保工作表范围足够大
            const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
            if (range.e.r < rowIndex) {
                range.e.r = rowIndex + 10; // 预留一些行
                sheet['!ref'] = XLSX.utils.encode_range(range);
            }

            // 插入新行
            XLSX.utils.sheet_add_aoa(sheet, [[]], {origin: rowIndex, insert: true});
        }

        // 复制行数据到新行
        function copyRowData(originalRowIndex, newRowIndex) {
            // 确保原始数据数组有足够的行
            while (originalData.length <= newRowIndex) {
                originalData.push([]);
            }

            // 复制原始数据到新行
            if (originalRowIndex < originalData.length) {
                const originalRow = originalData[originalRowIndex];
                const newRow = originalData[newRowIndex];

                // 复制所有列数据（除了电话号码列）
                for (let i = 0; i < originalRow.length; i++) {
                    if (i !== phoneColumn) {
                        newRow[i] = originalRow[i];
                    }
                }

                // 确保新行有足够的列
                while (newRow.length <= operatorColumn) {
                    newRow.push('');
                }
            }
        }

        function cleanPhoneNumber(phone) {
            return phone.replace(/\D/g, '');
        }

        async function getPhoneInfo(phoneNumber) {
            try {
                const response = await fetch(`${localProxyUrl}?number=${phoneNumber}`, {
                    method: 'GET',
                    timeout: 5000
                });

                if (!response.ok) {
                    throw new Error(`服务错误: ${response.status}`);
                }

                const data = await response.json();

                if (data.code === 0 && data.data) {
                    return {
                        location: `${data.data.province || ''}${data.data.city || ''}`.trim() || '未知地区',
                        operator: data.data.sp || '未知运营商'
                    };
                } else {
                    throw new Error('查询无结果');
                }
            } catch (error) {
                if (error.message.includes('Failed to fetch')) {
                    throw new Error('服务连接失败');
                } else if (error.message.includes('timeout')) {
                    throw new Error('请求超时');
                } else {
                    throw new Error(error.message);
                }
            }
        }

        function downloadResultFile() {
            if (!processedWorkbook || !selectedFile) return;

            // 生成新文件名
            const originalName = selectedFile.name;
            const nameWithoutExt = originalName.substring(0, originalName.lastIndexOf('.'));
            const ext = originalName.substring(originalName.lastIndexOf('.'));
            const newFileName = `${nameWithoutExt}_已查询${ext}`;

            // 转换为Excel文件并下载
            const excelBuffer = XLSX.write(processedWorkbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);

            const a = document.createElement('a');
            a.href = url;
            a.download = newFileName;
            a.click();

            // 释放URL资源
            setTimeout(() => URL.revokeObjectURL(url), 100);
        }

        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        function letterToColumnIndex(letter) {
            let index = 0;
            for (let i = 0; i < letter.length; i++) {
                index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
            }
            return index - 1;
        }

        function columnIndexToLetter(index) {
            let letter = '';
            let remaining = index;

            while (remaining >= 0) {
                const charCode = (remaining % 26) + 65; // 65 is 'A'
                letter = String.fromCharCode(charCode) + letter;
                remaining = Math.floor(remaining / 26) - 1;
            }

            return letter;
        }

        function addLog(message) {
            const logEntry = document.createElement('div');
            logEntry.className = 'py-1 border-b border-gray-100 last:border-0';
            logEntry.textContent = message;
            recentLog.appendChild(logEntry);
            recentLog.scrollTop = recentLog.scrollHeight;
        }

        document.addEventListener('DOMContentLoaded', initEventListeners);
    </script>
</body>
</html>"""

# 后端服务逻辑
app = Flask(__name__)
CORS(app, resources={r"/query": {"origins": "*"}})  # 允许所有本地请求
API_URL = "https://cx.shouji.360.cn/phonearea.php"


@app.route('/query', methods=['GET'])
def query_phone():
    """处理手机号查询请求"""
    phone_number = request.args.get('number', '')

    if not phone_number or not phone_number.isdigit() or len(phone_number) != 11:
        return jsonify({"code": -1, "msg": "无效的手机号"}), 400

    try:
        response = requests.get(API_URL, params={"number": phone_number}, timeout=5)
        response.raise_for_status()
        return jsonify(response.json())

    except requests.exceptions.RequestException as e:
        return jsonify({"code": -2, "msg": f"查询失败: {str(e)}"}), 500


@app.route('/')
def serve_frontend():
    """提供前端页面"""
    # 替换前端中的代理地址为实际地址
    proxy_url = f"http://localhost:{app.config['PORT']}/query"
    return HTML_CONTENT.replace('__PROXY_URL__', proxy_url)


def find_available_port(start=5000, max_attempts=20):
    """查找可用端口"""
    for port in range(start, start + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('localhost', port))
                return port
            except OSError:
                continue
    return None


def run_server(port):
    """启动Flask服务"""
    app.config['PORT'] = port
    app.run(host='localhost', port=port, debug=False, use_reloader=False)


def main():
    # 查找可用端口
    port = find_available_port()
    if not port:
        print("错误：无法找到可用端口，请关闭占用端口的程序后重试")
        input("按任意键退出...")
        return

    # 启动后端服务（在后台线程）
    server_thread = threading.Thread(target=run_server, args=(port,), daemon=True)
    server_thread.start()

    # 等待服务启动
    time.sleep(1)

    # 打开浏览器
    print(f"服务已启动，端口: {port}")
    print("正在打开浏览器...")
    webbrowser.open(f"http://localhost:{port}")

    # 保持主线程运行
    print("应用运行中，关闭窗口将终止服务")
    try:
        while True:
            time.sleep(3600)  # 休眠1小时
    except KeyboardInterrupt:
        print("服务已停止")


if __name__ == '__main__':
    main()
