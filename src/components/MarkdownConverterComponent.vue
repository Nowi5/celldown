<template>
  <div>
    <!-- Input area -->
    <div class="field">
        <label class="label">Input</label>
        <div class="control">
            <textarea class="textarea is-fullwidth" v-model="userInput" @input="processInput" placeholder="Paste Excel data..."></textarea>
        </div>
    </div>

    <!-- Output area -->    
    <div v-if="markdownOutput">
        <!-- Display the Markdown Output -->
        <!-- <pre>{{ markdownOutput }}</pre> -->

        <div class="field is-grouped">
                <div class="columns is-vcentered is-mobile is-gapless">
                    <!-- Label -->
                    <div class="column is-narrow">
                        <label class="label is-small">Sort table:</label>
                    </div>

                    <!-- Dropdown for columns -->
                    <div class="column is-narrow">
                        <div class="select is-small">
                            <select v-model="selectedColumn">
                                <option v-for="(colName, index) in this.jsonData[0]" :value="index" :key="index">{{ colName }}</option>
                            </select>
                        </div>
                    </div>

                    <!-- Dropdown for sort direction -->
                    <div class="column is-narrow">
                        <div class="select is-small">
                            <select v-model="sortDirection">
                                <option value="asc">Ascending</option>
                                <option value="desc">Descending</option>
                            </select>
                        </div>
                    </div>

                    <!-- Button to trigger sorting -->
                    <div class="column is-narrow">
                        <label class="label is-small is-hidden">Sort Button</label> <!-- Hidden label for accessibility -->
                        <button class="button is-small is-light" @click="sortTable">Sort</button>
                    </div>
                </div>
            </div>

        <div class="field">
            <div class="control">
                <textarea class="textarea is-fullwidth" id="markdownOutput" v-model="markdownOutput" readonly placeholder="Markdown output will appear here..."></textarea>
            </div>
        </div>

        <button class="button m-3 is-primary" @click="copyToClipboard">{{ buttonText }}</button>

    </div>
    
  </div>
</template>

<script>
import * as XLSX from 'xlsx';
//import TurndownService from 'turndown';

export default {
  data() {
    return {
        userInput: '',
        markdownOutput: '',
        jsonData: [],
        selectedColumn: 0,       // default to first column
        sortDirection: 'asc',     // default to ascending
        buttonText: 'Copy to Clipboard'
    };
  },
  methods: {
    generateMarkdownTable() {
        // Convert JSON data to Markdown table format
        let markdownTable = "| ";
        
        // Header
        this.jsonData[0].forEach(header => {
        markdownTable += `${header} | `;
        });
        markdownTable += "\n| ";

        // Determine column widths
        const columnWidths = this.jsonData[0].map((header, colIndex) => {
        let maxLength = header.length;
        for (let rowIndex = 1; rowIndex < this.jsonData.length; rowIndex++) {
            const cellValue = this.jsonData[rowIndex][colIndex];
            if (cellValue && cellValue.toString().length > maxLength) {
            maxLength = cellValue.toString().length;
            }
        }
        return maxLength;
        });

        // Divider with adjusted width
        columnWidths.forEach(width => {
        markdownTable += `${'-'.repeat(width)} | `;
        });
        markdownTable += "\n";

        // Data Rows with adjusted width
        for (let i = 1; i < this.jsonData.length; i++) {
        markdownTable += "| ";
        this.jsonData[i].forEach((cell, colIndex) => {
            const spacesToAdd = columnWidths[colIndex] - cell.toString().length;
            markdownTable += `${cell} ${' '.repeat(spacesToAdd)} | `;
        });
        markdownTable += "\n";
        }

        this.markdownOutput = markdownTable;
    },
    processInput() {
        // Convert Excel to a JSON object
        const workbook = XLSX.read(this.userInput, { type: 'string' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        this.jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (this.jsonData.length) {
            this.generateMarkdownTable();
        }
    },
    sortTable() {
        const header = this.jsonData[0];  // Extract the header row
        const rows = this.jsonData.slice(1);  // Extract data rows

        // Sort the data rows
        rows.sort((a, b) => {
            const valA = a[this.selectedColumn];
            const valB = b[this.selectedColumn];
            if (this.sortDirection === 'asc') {
                return valA > valB ? 1 : (valA < valB ? -1 : 0);
            } else {
                return valA < valB ? 1 : (valA > valB ? -1 : 0);
            }
        });

        // Merge the header and sorted rows back into jsonData
        this.jsonData = [header, ...rows];

        this.generateMarkdownTable();
    },
    copyToClipboard: function() {
            let textarea = document.getElementById('markdownOutput');
            textarea.select();
            document.execCommand('copy');

            // Update the button text and revert it after 3 seconds
            this.buttonText = 'Copied!';
            setTimeout(() => {
                this.buttonText = 'Copy to Clipboard';
            }, 3000);
    }
  }
};
</script>
