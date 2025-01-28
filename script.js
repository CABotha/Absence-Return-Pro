document.addEventListener('DOMContentLoaded', function() {
    // Initialize form elements
    const form = document.getElementById('returnToWorkForm');
    const excelBtn = document.getElementById('exportExcel');
    const pdfBtn = document.getElementById('exportPdf');
    
    // Initially disable export buttons
    excelBtn.disabled = true;
    pdfBtn.disabled = true;
    excelBtn.style.opacity = '0.5';
    pdfBtn.style.opacity = '0.5';

    // Store uploaded files data
    let uploadedFilesData = [];

    // Get current date formatted as YYYY-MM-DD
    function getFormattedDate() {
        return new Date().toISOString().split('T')[0];
    }

    // Setup conditional fields
    function setupConditionalFields() {
        const conditionalFields = [
            {
                trigger: 'absenceType',
                target: 'otherAbsenceType',
                condition: value => value === 'other'
            },
            {
                trigger: 'fullyRecovered',
                target: 'notFullyRecoveredDetails',
                condition: value => value === 'no' || value === 'unsure'
            },
            {
                trigger: 'ongoingHealth',
                target: 'healthDetails',
                condition: value => value === 'yes' || value === 'unsure'
            },
            {
                trigger: 'likelyRecur',
                target: 'recurDetails',
                condition: value => value === 'yes'
            },
            {
                trigger: 'healthcareRecommendations',
                target: 'recommendationDetails',
                condition: value => value === 'yes'
            },
            {
                trigger: 'supportRequired',
                target: 'supportDetails',
                condition: value => value === 'yes'
            },
            {
                trigger: 'workFactors',
                target: 'workFactorsDetails',
                condition: value => value === 'yes' || value === 'unsure'
            },
            {
                trigger: 'returnConcerns',
                target: 'concernDetails',
                condition: value => value === 'yes' || value === 'unsure'
            },
            {
                trigger: 'documentsReceived',
                target: 'documentDetails',
                condition: value => value === 'yes'
            }
        ];

        conditionalFields.forEach(field => {
            const triggerElement = document.getElementById(field.trigger);
            const targetElement = document.getElementById(field.target);

            if (triggerElement && targetElement) {
                triggerElement.addEventListener('change', function() {
                    const shouldShow = field.condition(this.value);
                    targetElement.classList.toggle('hidden', !shouldShow);
                    targetElement.required = shouldShow;
                });
            }
        });
    }

    // Handle file uploads
    function setupFileUpload() {
        const fileInput = document.getElementById('fileUpload');
        const uploadedFiles = document.getElementById('uploadedFiles');

        if (fileInput) {
            fileInput.addEventListener('change', async function(e) {
                const files = Array.from(e.target.files);
                
                for (const file of files) {
                    try {
                        const fileData = await readFileAsBase64(file);
                        const fileItem = document.createElement('div');
                        fileItem.className = 'file-item';
                        
                        // Create preview if it's an image
                        let preview = '';
                        if (file.type.startsWith('image/')) {
                            preview = `<img src="${fileData}" style="max-width: 100px; max-height: 100px;">`;
                        }

                        fileItem.innerHTML = `
                            <div class="file-info">
                                ${preview}
                                <span>${file.name}</span>
                            </div>
                            <button type="button" class="remove-file">Remove</button>
                        `;

                        // Store file data
                        uploadedFilesData.push({
                            name: file.name,
                            type: file.type,
                            data: fileData,
                            preview: preview
                        });

                        uploadedFiles.appendChild(fileItem);
                    } catch (error) {
                        console.error('Error reading file:', error);
                    }
                }
            });

            uploadedFiles.addEventListener('click', function(e) {
                if (e.target.classList.contains('remove-file')) {
                    const fileItem = e.target.parentElement;
                    const fileName = fileItem.querySelector('span').textContent;
                    uploadedFilesData = uploadedFilesData.filter(f => f.name !== fileName);
                    fileItem.remove();
                }
            });
        }
    }

    // Read file as base64
    function readFileAsBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => reject(reader.error);
            reader.readAsDataURL(file);
        });
    }

    // Collect form data
    function getFormData() {
        return {
            "Basic Information": {
                "Employee Name": document.getElementById('employeeName').value || '-',
                "Employee ID": document.getElementById('employeeId').value || '-',
                "Interviewer Name": document.getElementById('interviewerName').value || '-',
                "Date of Interview": document.getElementById('interviewDate').value || '-'
            },
            "Absence Details": {
                "Start Date of Absence": document.getElementById('absenceStartDate').value || '-',
                "Return to Work Date": document.getElementById('returnDate').value || '-',
                "Type of Absence": document.getElementById('absenceType').value || '-',
                "Other Absence Type": document.getElementById('otherAbsenceType').value || '-',
                "Reason for Absence": document.getElementById('absenceReason').value || '-'
            },
            "Health and Well-being": {
                "Fully Recovered": document.getElementById('fullyRecovered').value || '-',
                "Recovery Details": document.getElementById('notFullyRecoveredDetails').value || '-',
                "Ongoing Health Issues": document.getElementById('ongoingHealth').value || '-',
                "Health Details": document.getElementById('healthDetails').value || '-',
                "Likely to Recur": document.getElementById('likelyRecur').value || '-',
                "Recurrence Details": document.getElementById('recurDetails').value || '-'
            },
            "Medical Treatment": {
                "Medical Treatment Received": document.getElementById('medicalTreatment').value || '-',
                "Healthcare Recommendations": document.getElementById('healthcareRecommendations').value || '-',
                "Recommendation Details": document.getElementById('recommendationDetails').value || '-'
            },
            "Support and Adjustments": {
                "Support Required": document.getElementById('supportRequired').value || '-',
                "Support Details": document.getElementById('supportDetails').value || '-',
                "Assistance Program": document.getElementById('assistanceProgram').value || '-'
            },
            "Work-Related Factors": {
                "Work Factors": document.getElementById('workFactors').value || '-',
                "Work Factors Details": document.getElementById('workFactorsDetails').value || '-',
                "Return Concerns": document.getElementById('returnConcerns').value || '-',
                "Concern Details": document.getElementById('concernDetails').value || '-'
            },
            "Policy Review": {
                "Familiar with Procedure": document.getElementById('procedureKnowledge').value || '-',
                "Need Policy Clarification": document.getElementById('policyClarification').value || '-'
            },
            "Action Plan": {
                "Action 1": document.getElementById('action1').value || '-',
                "Action 2": document.getElementById('action2').value || '-',
                "Action 3": document.getElementById('action3').value || '-'
            },
            "Documentation": {
                "Documents Received": document.getElementById('documentsReceived').value || '-',
                "Document Details": document.getElementById('documentDetails').value || '-'
            },
            "Additional Comments": {
                "Employee Comments": document.getElementById('employeeComments').value || '-',
                "Interviewer Comments": document.getElementById('interviewerComments').value || '-'
            },
            "Sign Off": {
                "Employee Signature": document.getElementById('employeeSignature').value || '-',
                "Employee Sign Date": document.getElementById('employeeSignatureDate').value || '-',
                "Interviewer Signature": document.getElementById('interviewerSignature').value || '-',
                "Interviewer Sign Date": document.getElementById('interviewerSignatureDate').value || '-',
                "Next Review Date": document.getElementById('nextReviewDate').value || '-'
            }
        };
    }

    // Export to Excel
    function exportToExcel() {
        try {
            const formData = getFormData();
            const wsData = [];
            
            // Convert nested object to flat array for Excel
            Object.entries(formData).forEach(([section, data]) => {
                wsData.push([section, '']); // Section header
                Object.entries(data).forEach(([key, value]) => {
                    wsData.push([key, value]);
                });
                wsData.push(['', '']); // Empty row between sections
            });

            // Add uploaded files section
            if (uploadedFilesData.length > 0) {
                wsData.push(['Uploaded Documents', '']);
                uploadedFilesData.forEach(file => {
                    wsData.push([file.name, file.type]);
                });
            }

            // Create workbook and worksheet
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(wsData);

            // Set column widths
            ws['!cols'] = [
                { wch: 40 }, // Column A width
                { wch: 60 }  // Column B width
            ];

            // Style headers
            for (let i = 0; i < wsData.length; i++) {
                if (wsData[i][0] && !wsData[i][1]) { // Section headers
                    const cell = XLSX.utils.encode_cell({ r: i, c: 0 });
                    if (!ws[cell]) continue;
                    
                    ws[cell].s = {
                        font: { bold: true, color: { rgb: "FFFFFF" } },
                        fill: { fgColor: { rgb: "2E5AAB" } },
                        alignment: { horizontal: "left" }
                    };
                }
            }

            // Add worksheet to workbook and save
            XLSX.utils.book_append_sheet(wb, ws, "Return to Work Interview");
            XLSX.writeFile(wb, `Return_to_Work_${getFormattedDate()}.xlsx`);
            
            alert('Excel file has been generated successfully!');
        } catch (error) {
            console.error('Excel export error:', error);
            alert('Error exporting to Excel. Please try again.');
        }
    }

    // Export to PDF
    async function exportToPDF() {
        try {
            const formData = getFormData();
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();

            // Set document properties
            doc.setProperties({
                title: 'Return to Work Interview Form',
                subject: 'Return to Work Interview',
                author: 'HR Department',
                keywords: 'return to work, interview, HR',
                creator: 'Return to Work System'
            });

            // Add title and form data
            let yPosition = 20;

            // Add title
            doc.setFontSize(16);
            doc.setFont(undefined, 'bold');
            doc.text('Return to Work Interview Form', 105, yPosition, { align: 'center' });
            yPosition += 15;

            // Add form data
            Object.entries(formData).forEach(([section, data]) => {
                // Add page if needed
                if (yPosition > 250) {
                    doc.addPage();
                    yPosition = 20;
                }

                doc.setFontSize(12);
                doc.setFont(undefined, 'bold');
                doc.setFillColor(46, 90, 171);
                doc.setTextColor(255, 255, 255);
                doc.rect(15, yPosition, 180, 8, 'F');
                doc.text(section, 20, yPosition + 6);
                yPosition += 15;

                doc.setFontSize(10);
                doc.setFont(undefined, 'normal');
                doc.setTextColor(0, 0, 0);

                Object.entries(data).forEach(([key, value]) => {
                    // Add page if needed
                    if (yPosition > 250) {
                        doc.addPage();
                        yPosition = 20;
                    }

                    doc.setFillColor(240, 240, 240);
                    doc.rect(15, yPosition, 180, 7, 'F');
                    doc.text(key, 20, yPosition + 5);
                    yPosition += 10;

                    const lines = doc.splitTextToSize(value.toString(), 170);
                    doc.text(lines, 20, yPosition);
                    yPosition += (lines.length * 7) + 5;
                });

                yPosition += 5;
            });

            // Add uploaded documents
            if (uploadedFilesData.length > 0) {
                doc.addPage();
                yPosition = 20;

                doc.setFontSize(14);
                doc.setFont(undefined, 'bold');
                doc.text('Uploaded Documents', 105, yPosition, { align: 'center' });
                yPosition += 15;

                for (const file of uploadedFilesData) {
                    if (yPosition > 250) {
                        doc.addPage();
                        yPosition = 20;
                    }

                    doc.setFontSize(12);
                    doc.text(file.name, 15, yPosition);
                    yPosition += 10;

                    if (file.type.startsWith('image/')) {
                        try {
                            doc.addImage(file.data, 'JPEG', 15, yPosition, 180, 100);
                            yPosition += 110;
                        } catch (error) {
                            console.error('Error adding image:', error);
                        }
                    } else {
                        const fileTypeText = file.type === 'application/pdf' ? 
                            '(PDF Document Attached)' : '(Document Attached)';
                        doc.text(fileTypeText, 15, yPosition);
                        yPosition += 10;
                    }
                }
            }

            // Add footer to all pages
            for (let i = 1; i <= doc.internal.getNumberOfPages(); i++) {
                doc.setPage(i);
                doc.setFontSize(8);
                doc.setTextColor(128);
                doc.text('Return to Work Interview Form', 15, doc.internal.pageSize.height - 10);
                doc.text(`Page ${i} of ${doc.internal.getNumberOfPages()}`, 
                    doc.internal.pageSize.width - 25, doc.internal.pageSize.height - 10);
            }

            // Save the PDF
            doc.save(`Return_to_Work_${getFormattedDate()}.pdf`);
            alert('PDF file has been generated successfully!');
        } catch (error) {
            console.error('PDF export error:', error);
            alert('Error exporting to PDF. Please try again.');
        }
    }

    // Form validation
    function validateForm() {
        const requiredFields = form.querySelectorAll('[required]');
        let isValid = true;

        requiredFields.forEach(field => {
            if (!field.value) {
                isValid = false;
                field.classList.add('error');
                // Scroll to first error
                if (isValid === false) {
                    field.scrollIntoView({ behavior: 'smooth', block: 'center' });
                }
            } else {
                field.classList.remove('error');
            }
        });

        return isValid;
    }

    // Initialize form
    function initializeForm() {
        setupConditionalFields();
        setupFileUpload();

        // Form submission handler
        if (form) {
            form.addEventListener('submit', function(e) {
                e.preventDefault();
                
                if (validateForm()) {
                    // Enable export buttons
                    excelBtn.disabled = false;
                    pdfBtn.disabled = false;
                    excelBtn.style.opacity = '1';
                    pdfBtn.style.opacity = '1';
                    
                    alert('Form submitted successfully! You can now export to PDF or Excel.');
                } else {
                    alert('Please fill in all required fields before submitting.');
                }
            });
        }

        // Export button handlers
        if (excelBtn) excelBtn.addEventListener('click', exportToExcel);
        if (pdfBtn) pdfBtn.addEventListener('click', exportToPDF);

        // Handle real-time validation
        form.querySelectorAll('input, select, textarea').forEach(element => {
            element.addEventListener('input', function() {
                if (this.hasAttribute('required')) {
                    if (!this.value) {
                        this.classList.add('error');
                    } else {
                        this.classList.remove('error');
                    }
                }
            });

            // Add change event for selects to handle conditional fields
            if (element.tagName === 'SELECT') {
                element.addEventListener('change', function() {
                    const conditionalId = this.dataset.conditional;
                    if (conditionalId) {
                        const conditionalElement = document.getElementById(conditionalId);
                        if (conditionalElement) {
                            const showConditional = this.value === 'yes' || 
                                                  this.value === 'no' || 
                                                  this.value === 'unsure';
                            conditionalElement.classList.toggle('hidden', !showConditional);
                            conditionalElement.required = showConditional;
                        }
                    }
                });
            }
        });

        // Handle form reset
        form.addEventListener('reset', function() {
            // Reset error states
            form.querySelectorAll('.error').forEach(element => {
                element.classList.remove('error');
            });

            // Reset file upload section
            uploadedFilesData = [];
            const uploadedFiles = document.getElementById('uploadedFiles');
            if (uploadedFiles) {
                uploadedFiles.innerHTML = '';
            }

            // Disable export buttons
            excelBtn.disabled = true;
            pdfBtn.disabled = true;
            excelBtn.style.opacity = '0.5';
            pdfBtn.style.opacity = '0.5';

            // Hide all conditional fields
            document.querySelectorAll('.hidden').forEach(element => {
                element.style.display = 'none';
                element.required = false;
            });
        });

        // Initialize date fields with today's date
        const dateFields = form.querySelectorAll('input[type="date"]');
        const today = getFormattedDate();
        dateFields.forEach(field => {
            if (!field.value) {
                field.value = today;
            }
        });
    }

    // Initialize everything
    initializeForm();
});