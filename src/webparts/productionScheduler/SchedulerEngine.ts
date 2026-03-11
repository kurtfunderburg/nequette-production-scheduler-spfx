import { ISchedulerState, SharePointDataService } from './SharePointDataService';

const SCHEDULER_STYLE = `
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

        /* Theme Variables */
        :root {
            --bg-main: #f8f9fa;
            --bg-panel: #ffffff;
            --bg-hover: #f1f5f9;
            --border: #e2e8f0;
            --text-main: #1e293b;
            --text-muted: #64748b;
            --primary: #0f172a;
            --btn-text: #ffffff;

            --sidebar-width: 500px;
            --h-header: 44px;
            --h-phase: 44px;
            --h-milestone: 44px;
            --h-deliverable: 34px;
        }

        [data-theme="dark"] {
            --bg-main: #0f172a;
            --bg-panel: #1e293b;
            --bg-hover: #334155;
            --border: #334155;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --primary: #38bdf8;
            --btn-text: #0f172a;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Inter', sans-serif;
        }

        body {
            background-color: var(--bg-main);
            color: var(--text-main);
            height: 100vh;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            transition: background-color 0.3s, color 0.3s;
        }

        /* Header & Controls */
        header {
            background: #0f172a;
            color: white;
            padding: 1rem 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .branding {
            display: flex;
            flex-direction: column;
            gap: 0.2rem;
        }

        .project-inputs {
            display: flex;
            gap: 1rem;
            align-items: baseline;
        }

        .input-title {
            background: transparent;
            border: none;
            color: white;
            font-size: 1.2rem;
            font-weight: 500;
            letter-spacing: 1px;
            text-transform: uppercase;
            outline: none;
            width: 300px;
            border-bottom: 1px dashed transparent;
            transition: border-color 0.2s;
        }

        .input-title:focus,
        .input-title:hover {
            border-bottom: 1px dashed #94a3b8;
        }

        .input-number {
            background: transparent;
            border: none;
            color: #94a3b8;
            font-size: 0.85rem;
            outline: none;
            width: 100px;
            border-bottom: 1px dashed transparent;
        }

        .input-number:focus,
        .input-number:hover {
            border-bottom: 1px dashed #64748b;
            color: white;
        }

        .controls {
            display: flex;
            gap: 1.5rem;
            align-items: center;
            background: var(--bg-panel);
            padding: 1rem 2rem;
            border-bottom: 1px solid var(--border);
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
            z-index: 10;
        }

        .control-group {
            display: flex;
            flex-direction: column;
            gap: 0.3rem;
        }

        .control-group label {
            font-size: 0.75rem;
            font-weight: 600;
            color: var(--text-muted);
            text-transform: uppercase;
        }

        .control-group input[type="date"] {
            padding: 0.4rem;
            border: 1px solid var(--border);
            border-radius: 4px;
            font-size: 0.9rem;
            outline: none;
            background: var(--bg-main);
            color: var(--text-main);
        }

        .phase-toggles {
            display: flex;
            gap: 0.5rem;
            flex-wrap: wrap;
        }

        .toggle-btn {
            padding: 0.4rem 0.8rem;
            border: 1px solid var(--border);
            background: var(--bg-main);
            color: var(--text-main);
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.8rem;
            transition: all 0.2s;
        }

        .toggle-btn.active {
            background: var(--primary);
            color: var(--btn-text);
            border-color: var(--primary);
        }

        .data-controls {
            margin-left: auto;
            display: flex;
            align-items: center;
            gap: 1rem;
        }

        .btn-view-toggle {
            background: var(--bg-panel);
            color: var(--primary);
            border: 1px solid var(--border);
            padding: 0.4rem 0.8rem;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.75rem;
            font-weight: 600;
            transition: all 0.2s;
        }

        .btn-view-toggle:hover {
            background: var(--bg-hover);
        }

        .btn-theme {
            background: var(--bg-panel);
            color: var(--primary);
            border: 1px solid var(--border);
            padding: 0.4rem 0.8rem;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.75rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 0.3rem;
        }

        .btn-theme:hover {
            background: var(--bg-hover);
        }

        .btn-export {
            background: #10b981;
            color: white;
            border: none;
            padding: 0.4rem 0.8rem;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.75rem;
            font-weight: 600;
        }

        .btn-clear {
            background: #ef4444;
            color: white;
            border: none;
            padding: 0.4rem 0.8rem;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.75rem;
        }

        /* Workspace Core */
        .workspace {
            display: flex;
            flex: 1;
            overflow: hidden;
            position: relative;
            background: var(--bg-panel);
        }

        /* Layout Wrappers */
        .gantt-wrapper {
            display: flex;
            flex: 1;
            width: 100%;
            height: 100%;
        }

        .calendar-wrapper {
            display: none;
            flex: 1;
            flex-direction: column;
            background: var(--bg-main);
            overflow: hidden;
        }

        /* Unified Gantt Layout */
        .gantt-master-container {
            flex: 1;
            overflow: auto;
            position: relative;
            background: var(--bg-panel);
            display: flex;
            flex-direction: column;
        }

        .unified-header-row {
            display: flex;
            position: sticky;
            top: 0;
            z-index: 30;
            width: fit-content;
            min-width: 100%;
            border-bottom: 1px solid var(--border);
            background: var(--bg-main);
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        }

        .header-left {
            width: var(--sidebar-width);
            flex-shrink: 0;
            position: sticky;
            left: 0;
            z-index: 31;
            background: var(--bg-main);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 1rem;
            border-right: 1px solid var(--border);
            font-weight: 600;
            font-size: 0.85rem;
            height: var(--h-header);
        }

        .header-right {
            display: flex;
            height: var(--h-header);
            background: var(--bg-main);
        }

        .gantt-month {
            border-right: 1px solid var(--border);
            text-align: center;
            font-size: 0.8rem;
            font-weight: 600;
            line-height: calc(var(--h-header) - 1px);
            box-sizing: border-box;
        }

        .gantt-rows-container {
            display: flex;
            flex-direction: column;
            width: fit-content;
            min-width: 100%;
            padding-bottom: 4rem;
            position: relative;
        }

        .unified-row {
            display: flex;
            width: 100%;
            border-bottom: 1px solid var(--border);
        }

        /* Row Left Sidebar */
        .row-left {
            width: var(--sidebar-width);
            flex-shrink: 0;
            position: sticky;
            left: 0;
            z-index: 20;
            background: var(--bg-panel);
            border-right: 1px solid var(--border);
        }

        .row-left.phase {
            background: var(--bg-hover);
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0 1rem;
            font-weight: 600;
            min-height: var(--h-phase);
        }

        .row-left.milestone {
            background: var(--bg-panel);
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0 1rem 0 2rem;
            min-height: var(--h-milestone);
        }

        .row-left.deliverable {
            background: var(--bg-panel);
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0 1rem 0 3.5rem;
            font-size: 0.8rem;
            color: var(--text-muted);
            min-height: var(--h-deliverable);
            border-bottom: 1px solid transparent;
        }

        /* Row Right Gantt Grid */
        .row-right {
            flex: 1;
            position: relative;
            background-image: linear-gradient(to right, var(--border) 1px, transparent 1px);
            background-size: 20px 100%;
        }

        [data-theme="dark"] .row-right.phase {
            background-color: rgba(51, 65, 85, 0.4);
        }

        [data-theme="light"] .row-right.phase {
            background-color: rgba(241, 245, 249, 0.6);
        }

        /* Draggable Bars */
        .bar-phase {
            position: absolute;
            height: 8px;
            top: 50%;
            transform: translateY(-50%);
            background: #64748b;
            border-radius: 4px;
            z-index: 5;
            cursor: grab;
            transition: box-shadow 0.2s;
            display: flex;
        }

        .bar-milestone {
            position: absolute;
            height: 24px;
            top: 50%;
            transform: translateY(-50%);
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            display: flex;
            align-items: center;
            font-size: 0.75rem;
            color: white;
            white-space: nowrap;
            overflow: hidden;
            z-index: 10;
            cursor: grab;
            transition: box-shadow 0.2s;
        }

        .bar-milestone-text {
            padding-left: 0.5rem;
            flex: 1;
            overflow: hidden;
            color: white;
        }

        .bar-deliverable {
            position: absolute;
            height: 12px;
            top: 50%;
            transform: translateY(-50%);
            border-radius: 3px;
            background: var(--border);
            border: 1px solid var(--text-muted);
            z-index: 5;
            cursor: grab;
            transition: box-shadow 0.2s;
            display: flex;
        }

        .bar-phase:active,
        .bar-milestone:active,
        .bar-deliverable:active {
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
            z-index: 50;
        }

        .resizer {
            position: absolute;
            right: 0;
            top: 0;
            bottom: 0;
            width: 6px;
            cursor: ew-resize;
            background: rgba(255, 255, 255, 0.25);
            border-left: 1px solid rgba(0, 0, 0, 0.1);
            z-index: 15;
            border-top-right-radius: inherit;
            border-bottom-right-radius: inherit;
        }

        .resizer:hover {
            background: rgba(255, 255, 255, 0.6);
        }

        /* Today Marker */
        .today-marker-wrapper {
            position: absolute;
            top: 0;
            bottom: 0;
            width: 2px;
            pointer-events: none;
            z-index: 25;
        }

        .today-line {
            position: absolute;
            top: 0;
            bottom: 0;
            width: 2px;
            background: repeating-linear-gradient(to bottom, #ef4444, #ef4444 6px, transparent 6px, transparent 12px);
        }

        .today-label {
            position: absolute;
            top: 4px;
            left: 50%;
            transform: translateX(-50%);
            background: #ef4444;
            color: white;
            font-size: 0.6rem;
            font-weight: 700;
            padding: 2px 6px;
            border-radius: 12px;
            white-space: nowrap;
        }

        /* UI Elements within Rows */
        .phase-header-left {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .phase-date-input {
            font-size: 0.75rem;
            padding: 0.2rem 0.4rem;
            border: 1px solid var(--border);
            border-radius: 4px;
            color: var(--text-main);
            outline: none;
            background: var(--bg-panel);
        }

        .phase-actions {
            display: flex;
            align-items: center;
            gap: 0.3rem;
        }

        .btn-add {
            background: none;
            border: none;
            font-size: 1.2rem;
            cursor: pointer;
            color: var(--text-muted);
            line-height: 1;
            display: flex;
            align-items: center;
            gap: 0.2rem;
        }

        .btn-add span {
            font-size: 0.65rem;
            text-transform: uppercase;
            font-weight: 600;
        }

        .btn-add:hover {
            color: var(--primary);
        }

        .btn-add-primary {
            background: var(--primary);
            color: var(--btn-text);
            border: none;
            padding: 0.3rem 0.6rem;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.7rem;
            text-transform: uppercase;
            font-weight: 600;
        }

        .milestone-info {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .color-dot {
            width: 12px;
            height: 12px;
            border-radius: 50%;
            border: 1px solid rgba(0, 0, 0, 0.1);
            flex-shrink: 0;
        }

        .milestone-name {
            font-size: 0.9rem;
            font-weight: 500;
            white-space: normal;
            line-height: 1.2;
            padding-top: 4px;
            padding-bottom: 4px;
            max-width: 250px;
        }

        .milestone-actions {
            display: flex;
            gap: 0.25rem;
            align-items: center;
        }

        .action-btn {
            font-size: 0.7rem;
            padding: 0.2rem 0.4rem;
            border: 1px solid var(--border);
            background: var(--bg-panel);
            color: var(--text-muted);
            cursor: pointer;
            border-radius: 3px;
        }

        .action-btn:hover:not(:disabled) {
            background: var(--bg-hover);
            color: var(--primary);
        }

        .action-btn:disabled {
            opacity: 0.3;
            cursor: not-allowed;
        }

        .deliverable-wrapper {
            display: flex;
            align-items: center;
            position: relative;
            width: 100%;
            justify-content: space-between;
        }

        .deliverable-wrapper::before {
            content: "└";
            position: absolute;
            left: -14px;
            color: var(--border);
        }

        .d-name {
            font-weight: 500;
            color: var(--text-main);
            white-space: normal;
            max-width: 200px;
        }

        .d-meta-actions {
            display: flex;
            align-items: center;
            gap: 0.25rem;
        }

        .d-meta {
            font-size: 0.7rem;
            background: var(--bg-hover);
            padding: 0.1rem 0.4rem;
            border-radius: 3px;
            border: 1px solid var(--border);
            margin-right: 4px;
        }


        /* Calendar View */
        .calendar-wrapper {
            display: none;
            flex: 1;
            flex-direction: column;
            background: var(--bg-main);
            overflow: hidden;
        }

        .cal-toolbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 1rem 2rem;
            background: var(--bg-panel);
            border-bottom: 1px solid var(--border);
            flex-shrink: 0;
        }

        .cal-nav {
            display: flex;
            align-items: center;
            gap: 1rem;
        }

        .btn-icon {
            background: var(--bg-panel);
            border: 1px solid var(--border);
            color: var(--text-main);
            border-radius: 4px;
            padding: 0.3rem 0.8rem;
            cursor: pointer;
            font-weight: 600;
            transition: background 0.2s;
        }

        .btn-icon:hover {
            background: var(--bg-hover);
        }

        .cal-title {
            font-size: 1.1rem;
            font-weight: 600;
            min-width: 220px;
            text-align: center;
            color: var(--primary);
        }

        .cal-actions {
            display: flex;
            align-items: center;
            gap: 1rem;
        }

        .view-toggles {
            display: flex;
            border: 1px solid var(--border);
            border-radius: 4px;
            overflow: hidden;
        }

        .view-toggles button {
            background: var(--bg-panel);
            border: none;
            padding: 0.4rem 1rem;
            font-size: 0.8rem;
            cursor: pointer;
            color: var(--text-muted);
            font-weight: 500;
        }

        .view-toggles button.active {
            background: var(--bg-hover);
            color: var(--primary);
            font-weight: 600;
            box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.02);
        }

        .cal-render-area {
            flex: 1;
            overflow-y: auto;
            padding: 2rem;
            background: var(--bg-main);
        }

        .cal-block {
            background: var(--bg-panel);
            border: 1px solid var(--border);
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
            display: flex;
            flex-direction: column;
            height: 100%;
        }

        .cal-days-row {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            background: var(--bg-main);
            border-bottom: 1px solid var(--border);
        }

        .cal-day-name {
            padding: 0.75rem 0.5rem;
            text-align: center;
            font-size: 0.75rem;
            font-weight: 600;
            color: var(--text-muted);
            text-transform: uppercase;
            border-right: 1px solid var(--border);
        }

        .cal-day-name:last-child {
            border-right: none;
        }

        .cal-grid {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            auto-rows: minmax(100px, 1fr);
            flex: 1;
        }

        .cal-mode-week .cal-grid {
            auto-rows: minmax(300px, 1fr);
        }

        .cal-cell {
            border-right: 1px solid var(--border);
            border-bottom: 1px solid var(--border);
            padding: 0.5rem;
            display: flex;
            flex-direction: column;
            gap: 0.2rem;
        }

        .cal-cell:nth-child(7n) {
            border-right: none;
        }

        .cal-cell.empty {
            background: var(--bg-main);
        }

        .cal-date {
            font-size: 0.8rem;
            font-weight: 600;
            margin-bottom: 0.5rem;
            color: var(--text-muted);
        }

        .cal-cell.today {
            background: rgba(239, 68, 68, 0.1);
        }

        .cal-cell.today .cal-date {
            color: #ef4444;
        }

        .cal-event {
            font-size: 0.7rem;
            padding: 0.25rem 0.4rem;
            border-radius: 3px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            color: white;
            display: flex;
            justify-content: space-between;
        }

        .cal-event.phase {
            background: var(--bg-hover);
            color: var(--text-main);
            border-left: 3px solid #64748b;
            font-weight: 600;
            margin-bottom: 2px;
        }

        .cal-event.milestone {
            border-left: 3px solid rgba(0, 0, 0, 0.2);
            margin-bottom: 1px;
        }

        .cal-event.deliverable {
            background: transparent;
            color: var(--text-main);
            border: 1px solid var(--border);
            border-left: 3px solid var(--text-muted);
        }

        /* Print Settings */
        @media print {
            @page {
                size: 17in 11in;
                margin: 0.5in;
            }

            body {
                background: white;
            }

            header,
            .controls,
            .gantt-master-container,
            .cal-toolbar {
                display: none !important;
            }

            .workspace,
            .calendar-wrapper {
                display: block;
                overflow: visible;
                padding: 0;
                background: white;
                height: auto;
            }

            .cal-render-area {
                padding: 0;
                overflow: visible;
                height: auto;
            }

            .cal-block {
                box-shadow: none;
                border: 2px solid #0f172a;
                border-radius: 0;
                margin: 0;
                break-inside: avoid;
                height: calc(100vh - 1in);
                background: white;
            }

            .cal-cell {
                border-color: #cbd5e1;
            }

            .cal-days-row {
                border-color: #0f172a;
                border-bottom: 2px solid #0f172a;
                background: white;
            }

            .cal-day-name {
                border-color: #cbd5e1;
                color: #0f172a;
            }

            .cal-grid {
                border-top: none;
            }

            .cal-date {
                color: #64748b;
            }

            * {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }

            .cal-block::before {
                content: attr(data-print-title);
                display: block;
                font-size: 1.5rem;
                font-weight: 700;
                color: #0f172a;
                padding: 0.5rem 1rem;
                background: #f1f5f9;
                border-bottom: 2px solid #0f172a;
                text-align: center;
                text-transform: uppercase;
                letter-spacing: 1px;
            }
        }

        /* Modals */
        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.6);
            justify-content: center;
            align-items: center;
            z-index: 1000;
        }

        .modal.active {
            display: flex;
        }

        .modal-content {
            background: var(--bg-panel);
            color: var(--text-main);
            padding: 2rem;
            border-radius: 6px;
            width: 450px;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.3);
        }

        .modal-content h3 {
            margin-bottom: 1rem;
            border-bottom: 1px solid var(--border);
            padding-bottom: 0.5rem;
        }

        .form-group {
            margin-bottom: 1rem;
            display: flex;
            flex-direction: column;
            gap: 0.3rem;
        }

        .form-group label {
            font-size: 0.8rem;
            font-weight: 500;
        }

        .form-group input {
            padding: 0.5rem;
            border: 1px solid var(--border);
            border-radius: 4px;
            background: var(--bg-main);
            color: var(--text-main);
        }

        .modal-actions {
            display: flex;
            justify-content: flex-end;
            gap: 0.5rem;
            margin-top: 1.5rem;
        }

        .btn-primary {
            background: var(--primary);
            color: var(--btn-text);
            padding: 0.5rem 1rem;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }

        .btn-secondary {
            background: var(--bg-main);
            color: var(--text-main);
            border: 1px solid var(--border);
            padding: 0.5rem 1rem;
            border-radius: 4px;
            cursor: pointer;
        }
    `;
const SCHEDULER_MARKUP = `

    <header>
        <div class="branding">
            <span style="font-size: 0.7rem; color: #94a3b8; text-transform: uppercase;">Nequette Architecture &amp;
                Design</span>
            <div class="project-inputs">
                <input type="text" id="project-name" class="input-title" placeholder="Project Name" title="Edit Project Name">
                <input type="text" id="project-number" class="input-number" placeholder="Proj #" title="Edit Project Number">
            </div>
        </div>
        <div class="data-controls">
            <span style="font-size: 0.75rem; color: #94a3b8;" id="helper-text">*Drag bars to move/resize.</span>
            <button class="btn-theme" onclick="toggleTheme()" title="Toggle Dark/Light Mode">◑ Theme</button>
            <button class="btn-view-toggle" id="view-toggle-btn" onclick="toggleViewMode()">Switch to Calendar
                View</button>
            <input type="file" id="json-import" accept=".json" style="display: none;" onchange="importData(event)">
            <button class="btn-export" style="background:#3b82f6;" onclick="document.getElementById('json-import').click()">⇪ Import JSON</button>
            <button class="btn-export" style="background:#3b82f6;" onclick="exportData()">⇩ Export JSON</button>
            <button class="btn-export" onclick="saveToHtmlFile()">⇩ Save HTML</button>
            <button class="btn-clear" onclick="clearData()">Reset</button>
        </div>
    </header>

    <div class="controls">
        <div class="control-group">
            <label>Complexity</label>
            <input type="number" id="complexity-multiplier" min="0.1" step="0.1" value="1" style="width: 80px; padding: 0.4rem; border: 1px solid var(--border); border-radius: 4px; font-size: 0.9rem; outline: none; background: var(--bg-main); color: var(--text-main);">
        </div>
        <div class="control-group">
            <label>Timeline Start Date</label>
            <input type="date" id="global-start">
        </div>
        <div class="control-group">
            <label>Phase Filters</label>
            <div class="phase-toggles" id="phase-toggles"><button class="toggle-btn active">Conceptual Design</button><button class="toggle-btn active">Schematic Design</button><button class="toggle-btn active">Design Development</button><button class="toggle-btn active">Copnstruction Documents</button></div>
        </div>
    </div>

    <div class="workspace">
        <div class="gantt-master-container" id="gantt-master-container">
            <div class="unified-header-row" id="unified-header-row">
                <div class="header-left">
                    <span>Data Hierarchy &amp; Deliverables</span>
                    <button class="btn-add-primary" onclick="openPhaseModal()">+ Phase</button>
                </div>
                <div class="header-right" id="gantt-header-right" style="min-width: 4800px;"><div class="gantt-month" style="width: 600px; flex-shrink: 0;">Mar 2026</div><div class="gantt-month" style="width: 600px; flex-shrink: 0;">Apr 2026</div><div class="gantt-month" style="width: 600px; flex-shrink: 0;">May 2026</div><div class="gantt-month" style="width: 600px; flex-shrink: 0;">Jun 2026</div><div class="gantt-month" style="width: 600px; flex-shrink: 0;">Jul 2026</div><div class="gantt-month" style="width: 600px; flex-shrink: 0;">Aug 2026</div><div class="gantt-month" style="width: 600px; flex-shrink: 0;">Sep 2026</div><div class="gantt-month" style="width: 600px; flex-shrink: 0;">Oct 2026</div></div>
            </div>

            <div class="gantt-rows-container" id="gantt-rows-container"><div class="unified-row"><div class="row-left phase">
                    <div class="phase-header-left">
                        <span style="max-width: 140px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Conceptual Design">Conceptual Design</span>
                        <input type="date" class="phase-date-input" value="2026-03-09" onchange="updatePhaseDate('cd', this.value)">
                    </div>
                    <div class="phase-actions">
                        <button class="action-btn" onclick="moveItem('phase', 'cd', null, null, -1)" disabled="">↑</button>
                        <button class="action-btn" onclick="moveItem('phase', 'cd', null, null, 1)">↓</button>
                        <button class="action-btn" onclick="openPhaseModal('cd')">E</button>
                        <button class="action-btn" onclick="deletePhase('cd')">X</button>
                        <button class="btn-add" onclick="openMilestoneModal('cd')"><span>+ MS</span></button>
                    </div>
                </div><div class="row-right phase" style="min-width: 4800px;"><div class="bar-phase" style="left: 20px; width: 300px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left milestone">
                        <div class="milestone-info"><div class="color-dot" style="background-color: #64748b"></div><span class="milestone-name">Massing &amp; Site</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', 'cd', 'm1', null, -1)" disabled="">↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', 'cd', 'm1', null, 1)" disabled="">↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('cd', 'm1')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('cd', 'm1')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('cd', 'm1')">X</button>
                        </div>
                    </div><div class="row-right" style="min-width: 4800px;"><div class="bar-milestone" style="left: 20px; width: 200px; background-color: rgb(100, 116, 139);"><span class="bar-milestone-text">Massing &amp; Site</span><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Site Plan</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">4d | <strong>Mar 13</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'cd', 'm1', 0, -1)" disabled="">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'cd', 'm1', 0, 1)">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('cd', 'm1', 0)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('cd', 'm1', 0)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 20px; width: 80px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Massing Model</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">10d | <strong>Mar 19</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'cd', 'm1', 1, -1)">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'cd', 'm1', 1, 1)" disabled="">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('cd', 'm1', 1)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('cd', 'm1', 1)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 20px; width: 200px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left phase">
                    <div class="phase-header-left">
                        <span style="max-width: 140px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Schematic Design">Schematic Design</span>
                        <input type="date" class="phase-date-input" value="2026-03-25" onchange="updatePhaseDate('sd', this.value)">
                    </div>
                    <div class="phase-actions">
                        <button class="action-btn" onclick="moveItem('phase', 'sd', null, null, -1)">↑</button>
                        <button class="action-btn" onclick="moveItem('phase', 'sd', null, null, 1)">↓</button>
                        <button class="action-btn" onclick="openPhaseModal('sd')">E</button>
                        <button class="action-btn" onclick="deletePhase('sd')">X</button>
                        <button class="btn-add" onclick="openMilestoneModal('sd')"><span>+ MS</span></button>
                    </div>
                </div><div class="row-right phase" style="min-width: 4800px;"><div class="bar-phase" style="left: 340px; width: 580px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left milestone">
                        <div class="milestone-info"><div class="color-dot" style="background-color: #0ea5e9"></div><span class="milestone-name">Revit Model Setup</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', 'sd', 'm2', null, -1)" disabled="">↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', 'sd', 'm2', null, 1)">↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('sd', 'm2')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('sd', 'm2')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('sd', 'm2')">X</button>
                        </div>
                    </div><div class="row-right" style="min-width: 4800px;"><div class="bar-milestone" style="left: 340px; width: 100px; background-color: rgb(14, 165, 233);"><span class="bar-milestone-text">Revit Model Setup</span><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Central File</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">2d | <strong>Mar 27</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm2', 0, -1)" disabled="">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm2', 0, 1)">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('sd', 'm2', 0)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('sd', 'm2', 0)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 340px; width: 40px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Worksets</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">5d | <strong>Mar 30</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm2', 1, -1)">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm2', 1, 1)" disabled="">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('sd', 'm2', 1)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('sd', 'm2', 1)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 340px; width: 100px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left milestone">
                        <div class="milestone-info"><div class="color-dot" style="background-color: #3b82f6"></div><span class="milestone-name">Core Layout</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', 'sd', 'm3', null, -1)">↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', 'sd', 'm3', null, 1)">↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('sd', 'm3')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('sd', 'm3')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('sd', 'm3')">X</button>
                        </div>
                    </div><div class="row-right" style="min-width: 4800px;"><div class="bar-milestone" style="left: 440px; width: 180px; background-color: rgb(59, 130, 246);"><span class="bar-milestone-text">Core Layout</span><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Floor Plans</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">4d | <strong>Apr 5</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm3', 0, -1)" disabled="">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm3', 0, 1)">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('sd', 'm3', 0)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('sd', 'm3', 0)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 480px; width: 80px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Elevations</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">3d | <strong>Apr 8</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm3', 1, -1)">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm3', 1, 1)" disabled="">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('sd', 'm3', 1)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('sd', 'm3', 1)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 560px; width: 60px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left milestone">
                        <div class="milestone-info"><div class="color-dot" style="background-color: #0162fe"></div><span class="milestone-name">50% SD review</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', 'sd', 'm_1772634756358', null, -1)">↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', 'sd', 'm_1772634756358', null, 1)">↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('sd', 'm_1772634756358')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('sd', 'm_1772634756358')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('sd', 'm_1772634756358')">X</button>
                        </div>
                    </div><div class="row-right" style="min-width: 4800px;"><div class="bar-milestone" style="left: 620px; width: 40px; background-color: rgb(1, 98, 254);"><span class="bar-milestone-text">50% SD review</span><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Design Check</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">1d | <strong>Apr 10</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm_1772634756358', 0, -1)" disabled="">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm_1772634756358', 0, 1)" disabled="">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('sd', 'm_1772634756358', 0)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('sd', 'm_1772634756358', 0)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 640px; width: 20px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left milestone">
                        <div class="milestone-info"><div class="color-dot" style="background-color: #1f365c"></div><span class="milestone-name">95% review</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', 'sd', 'm_1772634946943', null, -1)">↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', 'sd', 'm_1772634946943', null, 1)" disabled="">↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('sd', 'm_1772634946943')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('sd', 'm_1772634946943')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('sd', 'm_1772634946943')">X</button>
                        </div>
                    </div><div class="row-right" style="min-width: 4800px;"><div class="bar-milestone" style="left: 840px; width: 240px; background-color: rgb(31, 54, 92);"><span class="bar-milestone-text">95% review</span><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Design Sign off</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">1d | <strong>Apr 21</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm_1772634946943', 0, -1)" disabled="">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'sd', 'm_1772634946943', 0, 1)" disabled="">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('sd', 'm_1772634946943', 0)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('sd', 'm_1772634946943', 0)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 860px; width: 20px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left phase">
                    <div class="phase-header-left">
                        <span style="max-width: 140px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Design Development">Design Development</span>
                        <input type="date" class="phase-date-input" value="2026-04-23" onchange="updatePhaseDate('dd', this.value)">
                    </div>
                    <div class="phase-actions">
                        <button class="action-btn" onclick="moveItem('phase', 'dd', null, null, -1)">↑</button>
                        <button class="action-btn" onclick="moveItem('phase', 'dd', null, null, 1)">↓</button>
                        <button class="action-btn" onclick="openPhaseModal('dd')">E</button>
                        <button class="action-btn" onclick="deletePhase('dd')">X</button>
                        <button class="btn-add" onclick="openMilestoneModal('dd')"><span>+ MS</span></button>
                    </div>
                </div><div class="row-right phase" style="min-width: 4800px;"><div class="bar-phase" style="left: 920px; width: 700px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left milestone">
                        <div class="milestone-info"><div class="color-dot" style="background-color: #479a57"></div><span class="milestone-name">50% DD review</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', 'dd', 'm_1772657146503', null, -1)" disabled="">↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', 'dd', 'm_1772657146503', null, 1)">↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('dd', 'm_1772657146503')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('dd', 'm_1772657146503')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('dd', 'm_1772657146503')">X</button>
                        </div>
                    </div><div class="row-right" style="min-width: 4800px;"><div class="bar-milestone" style="left: 920px; width: 320px; background-color: rgb(71, 154, 87);"><span class="bar-milestone-text">50% DD review</span><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">DD Kickoff Meeting</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">1d | <strong>Apr 24</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657146503', 0, -1)" disabled="">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657146503', 0, 1)">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('dd', 'm_1772657146503', 0)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('dd', 'm_1772657146503', 0)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 920px; width: 20px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Plans Coments</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">3d | <strong>Apr 29</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657146503', 1, -1)">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657146503', 1, 1)">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('dd', 'm_1772657146503', 1)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('dd', 'm_1772657146503', 1)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 980px; width: 60px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Elevation Coments</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">4d | <strong>Apr 30</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657146503', 2, -1)">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657146503', 2, 1)" disabled="">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('dd', 'm_1772657146503', 2)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('dd', 'm_1772657146503', 2)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 980px; width: 80px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left milestone">
                        <div class="milestone-info"><div class="color-dot" style="background-color: #4d8f58"></div><span class="milestone-name">95% DD review</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', 'dd', 'm_1772657243205', null, -1)">↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', 'dd', 'm_1772657243205', null, 1)" disabled="">↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('dd', 'm_1772657243205')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('dd', 'm_1772657243205')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('dd', 'm_1772657243205')">X</button>
                        </div>
                    </div><div class="row-right" style="min-width: 4800px;"><div class="bar-milestone" style="left: 1240px; width: 380px; background-color: rgb(77, 143, 88);"><span class="bar-milestone-text">95% DD review</span><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Base Model Lock</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">5d | <strong>May 14</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657243205', 0, -1)" disabled="">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657243205', 0, 1)">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('dd', 'm_1772657243205', 0)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('dd', 'm_1772657243205', 0)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 1240px; width: 100px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Engineer Backgrounds</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">2d | <strong>May 15</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657243205', 1, -1)">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657243205', 1, 1)">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('dd', 'm_1772657243205', 1)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('dd', 'm_1772657243205', 1)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 1320px; width: 40px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Design review</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">2d | <strong>May 26</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657243205', 2, -1)">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'dd', 'm_1772657243205', 2, 1)" disabled="">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('dd', 'm_1772657243205', 2)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('dd', 'm_1772657243205', 2)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 1540px; width: 40px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left phase">
                    <div class="phase-header-left">
                        <span style="max-width: 140px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="Copnstruction Documents">Copnstruction Documents</span>
                        <input type="date" class="phase-date-input" value="2026-05-11" onchange="updatePhaseDate('p_1772578767069', this.value)">
                    </div>
                    <div class="phase-actions">
                        <button class="action-btn" onclick="moveItem('phase', 'p_1772578767069', null, null, -1)">↑</button>
                        <button class="action-btn" onclick="moveItem('phase', 'p_1772578767069', null, null, 1)" disabled="">↓</button>
                        <button class="action-btn" onclick="openPhaseModal('p_1772578767069')">E</button>
                        <button class="action-btn" onclick="deletePhase('p_1772578767069')">X</button>
                        <button class="btn-add" onclick="openMilestoneModal('p_1772578767069')"><span>+ MS</span></button>
                    </div>
                </div><div class="row-right phase" style="min-width: 4800px;"><div class="bar-phase" style="left: 1280px; width: 1280px;"><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left milestone">
                        <div class="milestone-info"><div class="color-dot" style="background-color: #8b9b3b"></div><span class="milestone-name">50% CD review</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', 'p_1772578767069', 'm_1772578784693', null, -1)" disabled="">↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', 'p_1772578767069', 'm_1772578784693', null, 1)" disabled="">↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('p_1772578767069', 'm_1772578784693')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('p_1772578767069', 'm_1772578784693')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('p_1772578767069', 'm_1772578784693')">X</button>
                        </div>
                    </div><div class="row-right" style="min-width: 4800px;"><div class="bar-milestone" style="left: 1580px; width: 140px; background-color: rgb(139, 155, 59);"><span class="bar-milestone-text">50% CD review</span><div class="resizer"></div></div></div></div><div class="unified-row"><div class="row-left deliverable">
                            <div class="deliverable-wrapper">
                                <span class="d-name">Design review</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">2d | <strong>May 28</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'p_1772578767069', 'm_1772578784693', 0, -1)" disabled="">↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', 'p_1772578767069', 'm_1772578784693', 0, 1)" disabled="">↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('p_1772578767069', 'm_1772578784693', 0)">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('p_1772578767069', 'm_1772578784693', 0)">X</button>
                                </div>
                            </div>
                        </div><div class="row-right" style="min-width: 4800px;"><div class="bar-deliverable" style="left: 1580px; width: 40px;"><div class="resizer"></div></div></div></div><div id="active-today-marker" class="today-marker-wrapper" style="left: calc(var(--sidebar-width) + 20px);"><div class="today-line"></div><div class="today-label">TODAY</div></div></div>
        </div>

        <div class="calendar-wrapper" id="view-calendar">
            <div class="cal-toolbar">
                <div class="cal-nav">
                    <button class="btn-icon" onclick="calPrev()">←</button>
                    <div class="cal-title" id="cal-title-display">March 2026</div>
                    <button class="btn-icon" onclick="calNext()">→</button>
                </div>
                <div class="cal-actions">
                    <div class="view-toggles">
                        <button id="btn-mode-month" onclick="setCalMode('month')" class="active">Month</button>
                        <button id="btn-mode-week" onclick="setCalMode('week')">Week</button>
                    </div>
                    <button class="btn-export" style="background:var(--primary); color:var(--btn-text);" onclick="window.print()">🖶 Print PDF (11x17)</button>
                </div>
            </div>
            <div class="cal-render-area" id="cal-render-area"></div>
        </div>
    </div>

    <div class="modal" id="phase-modal">
        <div class="modal-content">
            <h3 id="p-modal-title">Add/Edit Phase</h3>
            <input type="hidden" id="p-id">
            <div class="form-group"><label>Phase Name</label><input type="text" id="p-name"></div>
            <div class="form-group"><label>Start Date</label><input type="date" id="p-start"></div>
            <div class="form-group"><label>Total Duration (Days)</label><input type="number" id="p-duration" min="1">
            </div>
            <div class="modal-actions">
                <button class="btn-secondary" onclick="closePhaseModal()">Cancel</button>
                <button class="btn-primary" onclick="savePhase()">Save Phase</button>
            </div>
        </div>
    </div>

    <div class="modal" id="milestone-modal">
        <div class="modal-content">
            <h3 id="modal-title">Add/Edit Milestone</h3>
            <input type="hidden" id="m-id"><input type="hidden" id="m-phase">
            <div class="form-group"><label>Milestone Name</label><input type="text" id="m-name"></div>
            <div class="form-group"><label>Start Offset (Days from Phase)</label><input type="number" id="m-start" min="0"></div>
            <div class="form-group"><label>Total Duration (Days)</label><input type="number" id="m-duration" min="1">
            </div>
            <div class="form-group"><label>Color Identifier</label><input type="color" id="m-color"></div>
            <div class="modal-actions">
                <button class="btn-secondary" onclick="closeMilestoneModal()">Cancel</button>
                <button class="btn-primary" onclick="saveMilestone()">Save</button>
            </div>
        </div>
    </div>

    <div class="modal" id="deliverable-modal">
        <div class="modal-content">
            <h3 id="d-modal-title">Add/Edit Deliverable</h3>
            <input type="hidden" id="d-phase"><input type="hidden" id="d-milestone"><input type="hidden" id="d-index">
            <div class="form-group"><label>Deliverable Name</label><input type="text" id="d-name"></div>
            <div class="form-group"><label>Duration (Days)</label><input type="number" id="d-duration" min="1"></div>
            <div class="modal-actions">
                <button class="btn-secondary" onclick="closeDeliverableModal()">Cancel</button>
                <button class="btn-primary" onclick="saveDeliverable()">Save</button>
            </div>
        </div>
    </div>

    `;
const SCHEDULER_SCRIPT = `
        const DAY_IN_MS = 86400000;
        const PIXELS_PER_DAY = 20;
        const STORAGE_KEY = 'nequette_schedule_data';
        const THEME_KEY = 'nequette_theme';

        let state = {};
        let calCurrentDate = new Date();
        let calRenderMode = 'month';

        function loadData() {
            const embeddedData = document.getElementById('app-data').textContent;
            let scriptState = JSON.parse(embeddedData);
            const spState = (window).__spfxSchedulerState || null;

            if (spState && spState.lastUpdated > (scriptState.lastUpdated || 0)) state = spState;
            else state = scriptState;

            const savedTheme = localStorage.getItem(THEME_KEY) || 'light';
            document.documentElement.setAttribute('data-theme', savedTheme);

            const [y, m, d] = state.globalStartDate.split('-');
            calCurrentDate = new Date(y, m - 1, d);
        }

        function saveData() {
            state.lastUpdated = Date.now();
            document.getElementById('app-data').textContent = JSON.stringify(state, null, 4);
            const saver = (window).__spfxSchedulerSave;
            if (saver) saver(state);
        }

        function saveToHtmlFile() {
            state.lastUpdated = Date.now();
            document.getElementById('app-data').textContent = JSON.stringify(state, null, 4);
            document.querySelectorAll('.modal').forEach(m => m.classList.remove('active'));

            const htmlContent = "<!DOCTYPE html>\\n" + document.documentElement.outerHTML;
            const blob = new Blob([htmlContent], { type: 'text/html' });
            const url = URL.createObjectURL(blob);

            const a = document.createElement('a');
            a.href = url;
            const safeName = (state.projectName || 'Schedule').replace(/[^a-z0-9]/gi, '_').toLowerCase();
            a.download = \`Nequette_\${state.projectNumber || '000'}_\${safeName}.html\`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        function exportData() {
            state.lastUpdated = Date.now();
            const dataStr = JSON.stringify(state, null, 4);
            const blob = new Blob([dataStr], { type: 'application/json' });
            const url = URL.createObjectURL(blob);

            const a = document.createElement('a');
            a.href = url;
            const safeName = (state.projectName || 'Schedule').replace(/[^a-z0-9]/gi, '_').toLowerCase();
            a.download = \`Nequette_\${state.projectNumber || '000'}_\${safeName}.json\`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        function importData(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    const importedState = JSON.parse(e.target.result);
                    if (importedState && typeof importedState === 'object' && importedState.phases) {
                        state = importedState;
                        saveData();
                        initDataBindings();

                        const [y, m, d] = state.globalStartDate.split('-');
                        calCurrentDate = new Date(y, m - 1, d);

                        document.getElementById('json-import').value = '';
                        render();
                        alert("Schedule imported successfully!");
                    } else {
                        alert("Invalid JSON format. Make sure you are importing a Nequette Schedule JSON file.");
                    }
                } catch (err) {
                    alert("Error parsing JSON file. The file may be corrupted.");
                    console.error("JSON parse error:", err);
                }
            };
            reader.readAsText(file);
        }

        function toggleTheme() {
            const current = document.documentElement.getAttribute('data-theme');
            const target = current === 'light' ? 'dark' : 'light';
            document.documentElement.setAttribute('data-theme', target);
            localStorage.setItem(THEME_KEY, target);
        }

        function clearData() {
            if (confirm("Are you sure you want to completely wipe this schedule?")) {
                state = JSON.parse(document.getElementById('app-data').textContent);
                state.lastUpdated = Date.now();
                saveData();
                const [y, m, d] = state.globalStartDate.split('-');
                calCurrentDate = new Date(y, m - 1, d);
                initDataBindings();
                render();
            }
        }

        function toggleViewMode() {
            state.viewMode = state.viewMode === 'gantt' ? 'calendar' : 'gantt';
            const btn = document.getElementById('view-toggle-btn');
            const helper = document.getElementById('helper-text');
            const gWrap = document.getElementById('gantt-master-container');
            const cWrap = document.getElementById('view-calendar');

            if (state.viewMode === 'gantt') {
                btn.innerText = "Switch to Calendar View";
                helper.style.display = 'inline';
                gWrap.style.display = 'flex';
                cWrap.style.display = 'none';
            } else {
                btn.innerText = "Switch to Gantt Editor";
                helper.style.display = 'none';
                gWrap.style.display = 'none';
                cWrap.style.display = 'flex';
            }
            saveData(); render();
        }

        function getMult() {
            return parseFloat(state.complexityMultiplier) || 1;
        }

        function initDataBindings() {
            document.getElementById('project-name').value = state.projectName || '';
            document.getElementById('project-number').value = state.projectNumber || '';
            document.getElementById('global-start').value = state.globalStartDate;
            document.getElementById('complexity-multiplier').value = state.complexityMultiplier || 1;

            if (state.viewMode === 'calendar') {
                document.getElementById('view-toggle-btn').innerText = "Switch to Gantt Editor";
                document.getElementById('helper-text').style.display = 'none';
                document.getElementById('gantt-master-container').style.display = 'none';
                document.getElementById('view-calendar').style.display = 'flex';
            }
        }

        // --- Hierarchy Movement Logic ---
        function moveItem(type, phaseId, milestoneId = null, index = null, direction) {
            if (type === 'phase') {
                const idx = state.phases.findIndex(p => p.id === phaseId);
                const target = idx + direction;
                if (target < 0 || target >= state.phases.length) return;
                const temp = state.phases[idx];
                state.phases[idx] = state.phases[target];
                state.phases[target] = temp;
            } else if (type === 'milestone') {
                const phase = state.phases.find(p => p.id === phaseId);
                const idx = phase.milestones.findIndex(m => m.id === milestoneId);
                const target = idx + direction;
                if (target < 0 || target >= phase.milestones.length) return;
                const temp = phase.milestones[idx];
                phase.milestones[idx] = phase.milestones[target];
                phase.milestones[target] = temp;
            } else if (type === 'deliverable') {
                const ms = state.phases.find(p => p.id === phaseId).milestones.find(m => m.id === milestoneId);
                const target = index + direction;
                if (target < 0 || target >= ms.deliverables.length) return;
                const temp = ms.deliverables[index];
                ms.deliverables[index] = ms.deliverables[target];
                ms.deliverables[target] = temp;
            }
            saveData(); render();
        }

        // --- Interaction Logic (Gantt) ---
        let dInfo = { active: false, mode: null, type: null, phaseId: null, milestoneId: null, deliverableIndex: null, element: null, startX: 0, initialVal: 0 };
        function onInteractionStart(e, mode, type, phaseId, milestoneId = null, deliverableIndex = null) {
            e.stopPropagation();
            dInfo = { active: true, mode, type, phaseId, milestoneId, deliverableIndex, element: mode === 'resize' ? e.target.parentElement : e.target, startX: e.clientX, initialVal: mode === 'resize' ? parseInt(e.target.parentElement.style.width, 10) || 0 : parseInt(e.target.style.left, 10) || 0 };
            document.addEventListener('mousemove', onInteractionMove);
            document.addEventListener('mouseup', onInteractionEnd);
        }
        function onInteractionMove(e) {
            if (!dInfo.active) return;
            let newVal = dInfo.initialVal + (e.clientX - dInfo.startX);
            if (dInfo.mode === 'drag') {
                if (newVal < 0) newVal = 0;
                dInfo.element.style.left = \`\${newVal}px\`;
            } else if (dInfo.mode === 'resize') {
                if (newVal < PIXELS_PER_DAY) newVal = PIXELS_PER_DAY;
                dInfo.element.style.width = \`\${newVal}px\`;
            }
        }
        function onInteractionEnd(e) {
            if (!dInfo.active) return;
            dInfo.active = false;
            document.removeEventListener('mousemove', onInteractionMove);
            document.removeEventListener('mouseup', onInteractionEnd);
            const deltaDays = Math.round((e.clientX - dInfo.startX) / PIXELS_PER_DAY);
            if (deltaDays !== 0) applyInteractionChanges(deltaDays); else render();
        }
        function applyInteractionChanges(deltaDays) {
            const phase = state.phases.find(p => p.id === dInfo.phaseId);
            const mult = getMult();
            if (dInfo.mode === 'drag') {
                if (dInfo.type === 'phase') {
                    const [y, m, d] = phase.startDate.split('-');
                    let dateObj = new Date(y, m - 1, d);
                    dateObj.setDate(dateObj.getDate() + deltaDays);
                    phase.startDate = dateObj.toISOString().split('T')[0];
                }
                else if (dInfo.type === 'milestone') {
                    const ms = phase.milestones.find(m => m.id === dInfo.milestoneId);
                    ms.startOffset = Math.max(0, ms.startOffset + (deltaDays / mult));
                }
                else if (dInfo.type === 'deliverable') {
                    const ms = phase.milestones.find(m => m.id === dInfo.milestoneId);
                    const deliv = ms.deliverables[dInfo.deliverableIndex];
                    deliv.startOffset = Math.max(0, (deliv.startOffset || 0) + (deltaDays / mult));
                }
            } else if (dInfo.mode === 'resize') {
                if (dInfo.type === 'phase') phase.duration = Math.max(1 / mult, (phase.duration || 1) + (deltaDays / mult));
                else if (dInfo.type === 'milestone') {
                    const ms = phase.milestones.find(m => m.id === dInfo.milestoneId);
                    ms.duration = Math.max(1 / mult, ms.duration + (deltaDays / mult));
                }
                else if (dInfo.type === 'deliverable') {
                    const ms = phase.milestones.find(m => m.id === dInfo.milestoneId);
                    const deliv = ms.deliverables[dInfo.deliverableIndex];
                    deliv.duration = Math.max(1 / mult, deliv.duration + (deltaDays / mult));
                }
            }
            saveData(); render();
        }

        function init() {
            loadData();
            initDataBindings();
            document.getElementById('project-name').addEventListener('change', (e) => { state.projectName = e.target.value; saveData(); render(); });
            document.getElementById('project-number').addEventListener('change', (e) => { state.projectNumber = e.target.value; saveData(); render(); });
            document.getElementById('global-start').addEventListener('change', (e) => { state.globalStartDate = e.target.value; saveData(); render(); });
            document.getElementById('complexity-multiplier').addEventListener('change', (e) => {
                let val = parseFloat(e.target.value);
                if (isNaN(val) || val <= 0) val = 1;
                state.complexityMultiplier = val;
                e.target.value = val;
                saveData();
                render();
            });

            render();
        }

        function render() {
            renderToggles();
            if (state.viewMode === 'gantt') {
                renderGanttScale();
                renderGanttLayout();
                renderTodayMarker();
            } else {
                renderCalendarLayout();
            }
        }

        function renderToggles() {
            const container = document.getElementById('phase-toggles');
            container.innerHTML = '';
            state.phases.forEach(phase => {
                const btn = document.createElement('button');
                btn.className = \`toggle-btn \${phase.visible ? 'active' : ''}\`;
                btn.innerText = phase.name;
                btn.onclick = () => { phase.visible = !phase.visible; saveData(); render(); };
                container.appendChild(btn);
            });
        }

        function parseDate(dateStr) {
            const [year, month, day] = dateStr.split('-');
            return new Date(year, month - 1, day);
        }

        function updatePhaseDate(phaseId, newDate) { const phase = state.phases.find(p => p.id === phaseId); if (phase) { phase.startDate = newDate; saveData(); render(); } }

        // --- UNIFIED GANTT ENGINE ---
        function renderGanttScale() {
            const header = document.getElementById('gantt-header-right');
            header.innerHTML = '';
            const startDate = parseDate(state.globalStartDate);

            const totalMonths = 8;
            header.style.minWidth = \`\${totalMonths * 30 * PIXELS_PER_DAY}px\`;

            for (let i = 0; i < totalMonths; i++) {
                const monthDiv = document.createElement('div');
                monthDiv.className = 'gantt-month';
                monthDiv.style.width = \`\${30 * PIXELS_PER_DAY}px\`;
                monthDiv.style.flexShrink = '0';

                let d = new Date(startDate.getFullYear(), startDate.getMonth() + i, 1);
                monthDiv.innerText = d.toLocaleString('default', { month: 'short', year: 'numeric' });
                header.appendChild(monthDiv);
            }
        }

        function renderGanttLayout() {
            const rowsContainer = document.getElementById('gantt-rows-container');
            rowsContainer.innerHTML = '';
            const globalStartObj = parseDate(state.globalStartDate);
            const mult = getMult();

            const totalWidth = 8 * 30 * PIXELS_PER_DAY;

            state.phases.forEach((phase, pIndex) => {
                if (!phase.visible) return;
                const phaseStartObj = parseDate(phase.startDate);
                let phaseOffsetDays = Math.round((phaseStartObj - globalStartObj) / DAY_IN_MS);
                if (phaseOffsetDays < 0) phaseOffsetDays = 0;

                let effPhaseDuration = (phase.duration || 15) * mult;

                const rowPhase = document.createElement('div');
                rowPhase.className = 'unified-row';

                const leftPhase = document.createElement('div');
                leftPhase.className = 'row-left phase';
                leftPhase.innerHTML = \`
                    <div class="phase-header-left">
                        <span style="max-width: 140px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="\${phase.name}">\${phase.name}</span>
                        <input type="date" class="phase-date-input" value="\${phase.startDate}" onchange="updatePhaseDate('\${phase.id}', this.value)">
                    </div>
                    <div class="phase-actions">
                        <button class="action-btn" onclick="moveItem('phase', '\${phase.id}', null, null, -1)" \${pIndex === 0 ? 'disabled' : ''}>↑</button>
                        <button class="action-btn" onclick="moveItem('phase', '\${phase.id}', null, null, 1)" \${pIndex === state.phases.length - 1 ? 'disabled' : ''}>↓</button>
                        <button class="action-btn" onclick="openPhaseModal('\${phase.id}')">E</button>
                        <button class="action-btn" onclick="deletePhase('\${phase.id}')">X</button>
                        <button class="btn-add" onclick="openMilestoneModal('\${phase.id}')"><span>+ MS</span></button>
                    </div>
                \`;

                const rightPhase = document.createElement('div');
                rightPhase.className = 'row-right phase';
                rightPhase.style.minWidth = \`\${totalWidth}px\`;

                const phaseBar = document.createElement('div');
                phaseBar.className = 'bar-phase';
                phaseBar.style.left = \`\${phaseOffsetDays * PIXELS_PER_DAY}px\`;
                phaseBar.style.width = \`\${effPhaseDuration * PIXELS_PER_DAY}px\`;
                phaseBar.onmousedown = (e) => onInteractionStart(e, 'drag', 'phase', phase.id);

                const pResizer = document.createElement('div');
                pResizer.className = 'resizer';
                pResizer.onmousedown = (e) => onInteractionStart(e, 'resize', 'phase', phase.id);
                phaseBar.appendChild(pResizer);

                rightPhase.appendChild(phaseBar);
                rowPhase.appendChild(leftPhase);
                rowPhase.appendChild(rightPhase);
                rowsContainer.appendChild(rowPhase);

                phase.milestones.forEach((m, mIndex) => {
                    let effMStartOffset = m.startOffset * mult;
                    let effMDuration = m.duration * mult;
                    let mStartDateObj = new Date(phaseStartObj.getTime() + (effMStartOffset * DAY_IN_MS));
                    const mTotalOffsetDays = phaseOffsetDays + effMStartOffset;

                    const rowMilestone = document.createElement('div');
                    rowMilestone.className = 'unified-row';

                    const leftMilestone = document.createElement('div');
                    leftMilestone.className = 'row-left milestone';
                    leftMilestone.innerHTML = \`
                        <div class="milestone-info"><div class="color-dot" style="background-color: \${m.color}"></div><span class="milestone-name">\${m.name}</span></div>
                        <div class="milestone-actions">
                            <button class="action-btn" onclick="moveItem('milestone', '\${phase.id}', '\${m.id}', null, -1)" \${mIndex === 0 ? 'disabled' : ''}>↑</button>
                            <button class="action-btn" onclick="moveItem('milestone', '\${phase.id}', '\${m.id}', null, 1)" \${mIndex === phase.milestones.length - 1 ? 'disabled' : ''}>↓</button>
                            <button class="action-btn" onclick="openDeliverableModal('\${phase.id}', '\${m.id}')">+</button>
                            <button class="action-btn" onclick="openMilestoneModal('\${phase.id}', '\${m.id}')">E</button>
                            <button class="action-btn" onclick="deleteMilestone('\${phase.id}', '\${m.id}')">X</button>
                        </div>
                    \`;

                    const rightMilestone = document.createElement('div');
                    rightMilestone.className = 'row-right';
                    rightMilestone.style.minWidth = \`\${totalWidth}px\`;

                    const mBar = document.createElement('div');
                    mBar.className = 'bar-milestone';
                    mBar.style.left = \`\${mTotalOffsetDays * PIXELS_PER_DAY}px\`;
                    mBar.style.width = \`\${effMDuration * PIXELS_PER_DAY}px\`;
                    mBar.style.backgroundColor = m.color;
                    mBar.onmousedown = (e) => onInteractionStart(e, 'drag', 'milestone', phase.id, m.id);
                    mBar.innerHTML = \`<span class="bar-milestone-text">\${m.name}</span>\`;

                    const mResizer = document.createElement('div');
                    mResizer.className = 'resizer';
                    mResizer.onmousedown = (e) => onInteractionStart(e, 'resize', 'milestone', phase.id, m.id);
                    mBar.appendChild(mResizer);

                    rightMilestone.appendChild(mBar);
                    rowMilestone.appendChild(leftMilestone);
                    rowMilestone.appendChild(rightMilestone);
                    rowsContainer.appendChild(rowMilestone);

                    m.deliverables.forEach((d, index) => {
                        let effDOffset = (d.startOffset || 0) * mult;
                        let effDDuration = d.duration * mult;
                        let dDueDate = new Date(mStartDateObj.getTime() + ((effDOffset + effDDuration) * DAY_IN_MS));

                        const rowDeliverable = document.createElement('div');
                        rowDeliverable.className = 'unified-row';

                        const leftDeliverable = document.createElement('div');
                        leftDeliverable.className = 'row-left deliverable';
                        leftDeliverable.innerHTML = \`
                            <div class="deliverable-wrapper">
                                <span class="d-name">\${d.name}</span>
                                <div class="d-meta-actions">
                                    <span class="d-meta">\${effDDuration}d | <strong>\${dDueDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}</strong></span>
                                    <button class="action-btn" onclick="moveItem('deliverable', '\${phase.id}', '\${m.id}', \${index}, -1)" \${index === 0 ? 'disabled' : ''}>↑</button>
                                    <button class="action-btn" onclick="moveItem('deliverable', '\${phase.id}', '\${m.id}', \${index}, 1)" \${index === m.deliverables.length - 1 ? 'disabled' : ''}>↓</button>
                                    <button class="action-btn" onclick="openDeliverableModal('\${phase.id}', '\${m.id}', \${index})">E</button>
                                    <button class="action-btn" onclick="deleteDeliverable('\${phase.id}', '\${m.id}', \${index})">X</button>
                                </div>
                            </div>
                        \`;

                        const rightDeliverable = document.createElement('div');
                        rightDeliverable.className = 'row-right';
                        rightDeliverable.style.minWidth = \`\${totalWidth}px\`;

                        const dBar = document.createElement('div');
                        dBar.className = 'bar-deliverable';
                        dBar.style.left = \`\${(mTotalOffsetDays + effDOffset) * PIXELS_PER_DAY}px\`;
                        dBar.style.width = \`\${effDDuration * PIXELS_PER_DAY}px\`;
                        dBar.onmousedown = (e) => onInteractionStart(e, 'drag', 'deliverable', phase.id, m.id, index);

                        const dResizer = document.createElement('div');
                        dResizer.className = 'resizer';
                        dResizer.onmousedown = (e) => onInteractionStart(e, 'resize', 'deliverable', phase.id, m.id, index);
                        dBar.appendChild(dResizer);

                        rightDeliverable.appendChild(dBar);
                        rowDeliverable.appendChild(leftDeliverable);
                        rowDeliverable.appendChild(rightDeliverable);
                        rowsContainer.appendChild(rowDeliverable);
                    });
                });
            });
        }

        function renderTodayMarker() {
            const existing = document.getElementById('active-today-marker');
            if (existing) existing.remove();

            const globalStartObj = parseDate(state.globalStartDate);
            const now = new Date();
            const todayObj = new Date(now.getFullYear(), now.getMonth(), now.getDate());
            const diffDays = Math.round((todayObj - globalStartObj) / DAY_IN_MS);

            if (diffDays >= 0) {
                const markerContainer = document.createElement('div');
                markerContainer.id = 'active-today-marker';
                markerContainer.className = 'today-marker-wrapper';
                markerContainer.style.left = \`calc(var(--sidebar-width) + \${diffDays * PIXELS_PER_DAY}px)\`;
                markerContainer.innerHTML = \`<div class="today-line"></div><div class="today-label">TODAY</div>\`;

                document.getElementById('gantt-rows-container').appendChild(markerContainer);
            }
        }

        // --- CALENDAR ENGINE ---
        function setCalMode(mode) {
            calRenderMode = mode;
            document.getElementById('btn-mode-month').classList.toggle('active', mode === 'month');
            document.getElementById('btn-mode-week').classList.toggle('active', mode === 'week');
            renderCalendarLayout();
        }

        function calPrev() {
            if (calRenderMode === 'month') calCurrentDate.setMonth(calCurrentDate.getMonth() - 1);
            else calCurrentDate.setDate(calCurrentDate.getDate() - 7);
            renderCalendarLayout();
        }

        function calNext() {
            if (calRenderMode === 'month') calCurrentDate.setMonth(calCurrentDate.getMonth() + 1);
            else calCurrentDate.setDate(calCurrentDate.getDate() + 7);
            renderCalendarLayout();
        }

        function renderCalendarLayout() {
            const renderArea = document.getElementById('cal-render-area');
            const titleDisplay = document.getElementById('cal-title-display');
            renderArea.innerHTML = '';
            renderArea.className = \`cal-render-area cal-mode-\${calRenderMode}\`;

            let flatEvents = [];
            const mult = getMult();
            state.phases.forEach(phase => {
                if (!phase.visible) return;
                const pStart = parseDate(phase.startDate);
                const pEnd = new Date(pStart.getTime() + ((phase.duration * mult) * DAY_IN_MS));
                flatEvents.push({ title: phase.name, start: pStart, end: pEnd, type: 'phase', color: null });

                phase.milestones.forEach(m => {
                    const mStart = new Date(pStart.getTime() + ((m.startOffset * mult) * DAY_IN_MS));
                    const mEnd = new Date(mStart.getTime() + ((m.duration * mult) * DAY_IN_MS));
                    flatEvents.push({ title: m.name, start: mStart, end: mEnd, type: 'milestone', color: m.color });

                    m.deliverables.forEach(d => {
                        const dStart = new Date(mStart.getTime() + (((d.startOffset || 0) * mult) * DAY_IN_MS));
                        const dEnd = new Date(dStart.getTime() + ((d.duration * mult) * DAY_IN_MS));
                        flatEvents.push({ title: d.name, start: dStart, end: dEnd, type: 'deliverable', color: null });
                    });
                });
            });

            const cYear = calCurrentDate.getFullYear();
            const cMonth = calCurrentDate.getMonth();
            const now = new Date();
            const todayStr = \`\${now.getFullYear()}-\${now.getMonth()}-\${now.getDate()}\`;

            let printTitleText = \`\${state.projectNumber} - \${state.projectName} - \`;

            if (calRenderMode === 'month') {
                titleDisplay.innerText = calCurrentDate.toLocaleString('default', { month: 'long', year: 'numeric' });
                printTitleText += titleDisplay.innerText;
                const monthBlock = createCalBlock(printTitleText);
                const grid = document.createElement('div'); grid.className = 'cal-grid';

                const firstDayIndex = new Date(cYear, cMonth, 1).getDay();
                for (let b = 0; b < firstDayIndex; b++) grid.appendChild(createEmptyCell());

                const daysInMonth = new Date(cYear, cMonth + 1, 0).getDate();
                for (let day = 1; day <= daysInMonth; day++) grid.appendChild(createDayCell(cYear, cMonth, day, todayStr, flatEvents));

                monthBlock.appendChild(grid); renderArea.appendChild(monthBlock);

            } else if (calRenderMode === 'week') {
                const startOfWeek = new Date(cYear, cMonth, calCurrentDate.getDate() - calCurrentDate.getDay());
                const endOfWeekStr = new Date(startOfWeek.getTime() + (6 * DAY_IN_MS)).toLocaleString('default', { month: 'short', day: 'numeric' });
                titleDisplay.innerText = \`Week of \${startOfWeek.toLocaleString('default', { month: 'short', day: 'numeric', year: 'numeric' })} - \${endOfWeekStr}\`;
                printTitleText += titleDisplay.innerText;

                const weekBlock = createCalBlock(printTitleText);
                const grid = document.createElement('div'); grid.className = 'cal-grid';

                for (let i = 0; i < 7; i++) {
                    const targetDay = new Date(startOfWeek.getTime() + (i * DAY_IN_MS));
                    grid.appendChild(createDayCell(targetDay.getFullYear(), targetDay.getMonth(), targetDay.getDate(), todayStr, flatEvents));
                }
                weekBlock.appendChild(grid); renderArea.appendChild(weekBlock);
            }
        }

        function createCalBlock(printDataTitle) {
            const block = document.createElement('div'); block.className = 'cal-block'; block.setAttribute('data-print-title', printDataTitle);
            const daysRow = document.createElement('div'); daysRow.className = 'cal-days-row';
            ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'].forEach(day => {
                const d = document.createElement('div'); d.className = 'cal-day-name'; d.innerText = day; daysRow.appendChild(d);
            });
            block.appendChild(daysRow); return block;
        }

        function createEmptyCell() { const cell = document.createElement('div'); cell.className = 'cal-cell empty'; return cell; }

        function createDayCell(year, month, day, todayStr, flatEvents) {
            const cell = document.createElement('div'); cell.className = 'cal-cell';
            if (\`\${year}-\${month}-\${day}\` === todayStr) cell.classList.add('today');

            const dateTag = document.createElement('div'); dateTag.className = 'cal-date'; dateTag.innerText = day; cell.appendChild(dateTag);
            const cellDateStart = new Date(year, month, day).getTime();
            const dayEvents = flatEvents.filter(ev => cellDateStart >= ev.start.getTime() && cellDateStart < ev.end.getTime());

            dayEvents.forEach(ev => {
                const eTag = document.createElement('div'); eTag.className = \`cal-event \${ev.type}\`;
                if (ev.type === 'milestone') eTag.style.backgroundColor = ev.color;
                let connector = cellDateStart > ev.start.getTime() ? '← ' : '';
                eTag.innerHTML = \`<span>\${connector}\${ev.title}</span>\`;
                cell.appendChild(eTag);
            });
            return cell;
        }

        // --- Modals (CRUD) ---
        function openPhaseModal(phaseId = null) {
            const pModal = document.getElementById('phase-modal');
            if (phaseId) {
                const p = state.phases.find(p => p.id === phaseId);
                document.getElementById('p-modal-title').innerText = "Edit Phase";
                document.getElementById('p-id').value = p.id; document.getElementById('p-name').value = p.name;
                document.getElementById('p-start').value = p.startDate; document.getElementById('p-duration').value = p.duration || 15;
            } else {
                document.getElementById('p-modal-title').innerText = "Add Phase";
                document.getElementById('p-id').value = ''; document.getElementById('p-name').value = '';
                document.getElementById('p-start').value = state.globalStartDate; document.getElementById('p-duration').value = '30';
            }
            pModal.classList.add('active');
        }
        function closePhaseModal() { document.getElementById('phase-modal').classList.remove('active'); }
        function savePhase() {
            const pId = document.getElementById('p-id').value;
            const newPhase = { id: pId || 'p_' + Date.now(), name: document.getElementById('p-name').value, startDate: document.getElementById('p-start').value, duration: parseInt(document.getElementById('p-duration').value, 10), visible: true, milestones: pId ? state.phases.find(p => p.id === pId).milestones : [] };
            if (pId) { const index = state.phases.findIndex(p => p.id === pId); state.phases[index] = newPhase; } else state.phases.push(newPhase);
            saveData(); closePhaseModal(); render();
        }
        function deletePhase(phaseId) { if (confirm("Delete this Phase?")) { state.phases = state.phases.filter(p => p.id !== phaseId); saveData(); render(); } }

        function openMilestoneModal(phaseId, milestoneId = null) {
            const mModal = document.getElementById('milestone-modal');
            document.getElementById('m-phase').value = phaseId; const phase = state.phases.find(p => p.id === phaseId);
            if (milestoneId) {
                document.getElementById('modal-title').innerText = "Edit Milestone"; const m = phase.milestones.find(m => m.id === milestoneId);
                document.getElementById('m-id').value = m.id; document.getElementById('m-name').value = m.name;
                document.getElementById('m-start').value = m.startOffset; document.getElementById('m-duration').value = m.duration; document.getElementById('m-color').value = m.color;
            } else {
                document.getElementById('modal-title').innerText = "Add Milestone"; document.getElementById('m-id').value = ''; document.getElementById('m-name').value = '';
                document.getElementById('m-start').value = '0'; document.getElementById('m-duration').value = '7'; document.getElementById('m-color').value = '#1e293b';
            }
            mModal.classList.add('active');
        }
        function closeMilestoneModal() { document.getElementById('milestone-modal').classList.remove('active'); }
        function saveMilestone() {
            const phaseId = document.getElementById('m-phase').value; const mId = document.getElementById('m-id').value; const phase = state.phases.find(p => p.id === phaseId);
            const newMilestone = { id: mId || 'm_' + Date.now(), name: document.getElementById('m-name').value, startOffset: parseInt(document.getElementById('m-start').value, 10), duration: parseInt(document.getElementById('m-duration').value, 10), color: document.getElementById('m-color').value, deliverables: mId ? phase.milestones.find(m => m.id === mId).deliverables : [] };
            if (mId) { const index = phase.milestones.findIndex(m => m.id === mId); phase.milestones[index] = newMilestone; } else phase.milestones.push(newMilestone);
            saveData(); closeMilestoneModal(); render();
        }
        function deleteMilestone(phaseId, milestoneId) { if (confirm("Delete this milestone?")) { const phase = state.phases.find(p => p.id === phaseId); phase.milestones = phase.milestones.filter(m => m.id !== milestoneId); saveData(); render(); } }

        function openDeliverableModal(phaseId, milestoneId, index = null) {
            const dModal = document.getElementById('deliverable-modal');
            document.getElementById('d-phase').value = phaseId; document.getElementById('d-milestone').value = milestoneId; document.getElementById('d-index').value = index !== null ? index : '';
            if (index !== null) {
                document.getElementById('d-modal-title').innerText = "Edit Deliverable"; const d = state.phases.find(p => p.id === phaseId).milestones.find(m => m.id === milestoneId).deliverables[index];
                document.getElementById('d-name').value = d.name; document.getElementById('d-duration').value = d.duration;
            } else {
                document.getElementById('d-modal-title').innerText = "Add Deliverable"; document.getElementById('d-name').value = ''; document.getElementById('d-duration').value = '5';
            }
            dModal.classList.add('active');
        }
        function closeDeliverableModal() { document.getElementById('deliverable-modal').classList.remove('active'); }
        function saveDeliverable() {
            const phaseId = document.getElementById('d-phase').value; const milestoneId = document.getElementById('d-milestone').value; const idx = document.getElementById('d-index').value;
            const ms = state.phases.find(p => p.id === phaseId).milestones.find(m => m.id === milestoneId);
            const newD = { name: document.getElementById('d-name').value, duration: parseInt(document.getElementById('d-duration').value, 10), startOffset: idx !== '' ? ms.deliverables[idx].startOffset : 0 };
            if (idx !== '') ms.deliverables[idx] = newD; else ms.deliverables.push(newD);
            saveData(); closeDeliverableModal(); render();
        }
        function deleteDeliverable(phaseId, milestoneId, index) { if (confirm("Delete this deliverable?")) { const ms = state.phases.find(p => p.id === phaseId).milestones.find(m => m.id === milestoneId); ms.deliverables.splice(index, 1); saveData(); render(); } }

        init();
    
        Object.assign(window, {
            saveToHtmlFile,
            exportData,
            importData,
            toggleTheme,
            clearData,
            toggleViewMode,
            setCalMode,
            calPrev,
            calNext,
            updatePhaseDate,
            moveItem,
            openPhaseModal,
            closePhaseModal,
            savePhase,
            deletePhase,
            openMilestoneModal,
            closeMilestoneModal,
            saveMilestone,
            deleteMilestone,
            openDeliverableModal,
            closeDeliverableModal,
            saveDeliverable,
            deleteDeliverable,
            onInteractionStart
        });
`;

export class SchedulerEngine {
  public constructor(private readonly host: HTMLElement, private readonly dataService: SharePointDataService) {}

  public async init(): Promise<void> {
    const loaded = await this.dataService.loadData();
    const state: ISchedulerState & { complexityMultiplier?: number } = {
      ...loaded,
      complexityMultiplier: (loaded as any).complexityMultiplier || 1
    };

    const windowAny = window as any;
    windowAny.__spfxSchedulerState = state;
    windowAny.__spfxSchedulerSave = (nextState: ISchedulerState): void => {
      this.dataService.saveData(nextState).catch((error: Error) => console.error('Scheduler save failed', error));
    };

    this.host.innerHTML = `
      <style>${SCHEDULER_STYLE}</style>
      <script id="app-data" type="application/json">${JSON.stringify(state, null, 4)}</script>
      ${SCHEDULER_MARKUP}
    `;

    const run = new Function(SCHEDULER_SCRIPT);
    run();
  }
}
