// Initialize AOS
AOS.init({
    duration: 1000,
    once: true,
    offset: 100
});

// Initialize Swiper
const swiper = new Swiper('.process-swiper', {
    slidesPerView: 1,
    spaceBetween: 30,
    loop: true,
    autoplay: {
        delay: 3000,
        disableOnInteraction: false,
    },
    pagination: {
        el: '.swiper-pagination',
        clickable: true,
    },
    navigation: {
        nextEl: '.swiper-button-next',
        prevEl: '.swiper-button-prev',
    },
    breakpoints: {
        640: {
            slidesPerView: 2,
        },
        768: {
            slidesPerView: 3,
        },
        1024: {
            slidesPerView: 4,
        },
    }
});

// Smooth scrolling for navigation links
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        // Only prevent default for actual anchor links, not blob URLs
        const href = this.getAttribute('href');
        if (href && href.startsWith('#')) {
            e.preventDefault();
            const target = document.querySelector(href);
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        }
    });
});

// File upload handling
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const targetLangSelect = document.getElementById('targetLangs');
const serviceSelect = document.getElementById('service');
const apiKeySection = document.getElementById('apiKeySection');

let targetLangChoices = null;
let defaultLangSelection = [];

// Load languages dynamically from API
async function loadLanguages() {
    try {
        const response = await fetch('/api/languages');
        const data = await response.json();
        const languages = data.languages;

        // Clear existing options
        targetLangSelect.innerHTML = '';

        // Popular languages first
        const popularLangs = ['es', 'fr', 'de', 'it', 'pt', 'ru', 'ja', 'ko', 'zh-CN', 'ar', 'hi'];

        // Add popular languages group
        const popularGroup = document.createElement('optgroup');
        popularGroup.label = '🌟 Popular Languages';
        popularLangs.forEach(code => {
            if (languages[code]) {
                const option = document.createElement('option');
                option.value = code;
                option.textContent = languages[code];
                if (code === 'es') option.selected = true;
                popularGroup.appendChild(option);
            }
        });
        targetLangSelect.appendChild(popularGroup);

        // Add all other languages
        const allGroup = document.createElement('optgroup');
        allGroup.label = '🌍 All Languages (A-Z)';
        Object.entries(languages).forEach(([code, name]) => {
            if (!popularLangs.includes(code)) {
                const option = document.createElement('option');
                option.value = code;
                option.textContent = name;
                allGroup.appendChild(option);
            }
        });
        targetLangSelect.appendChild(allGroup);

        // Initialize Choices.js after loading
        if (typeof Choices !== 'undefined') {
            targetLangChoices = new Choices(targetLangSelect, {
                removeItemButton: true,
                searchEnabled: true,
                placeholder: true,
                placeholderValue: 'Search and select languages...',
                shouldSort: false,
                itemSelectText: '',
                searchPlaceholderValue: 'Type to search...',
                noResultsText: 'No languages found',
                maxItemCount: 10
            });
            defaultLangSelection = targetLangChoices.getValue(true) || [];
        }
    } catch (error) {
        console.error('Failed to load languages:', error);
    }
}

// Load languages on page load
if (targetLangSelect) {
    loadLanguages();
}

// Drag and drop
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');

    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.pptx')) {
        handleFileSelect(files[0]);
    } else {
        Swal.fire({
            icon: 'warning',
            title: 'Invalid File',
            text: 'Please select a valid PowerPoint file (.pptx)',
            confirmButtonColor: '#1E40AF'
        });
    }
});

// File input change
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

// Handle file selection
function handleFileSelect(file) {
    if (file.name.endsWith('.pptx')) {
        fileName.textContent = file.name;

        // Calculate and display file size in MB
        const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
        const fileSize = document.getElementById('fileSize');
        fileSize.textContent = `${fileSizeMB} MB`;

        document.querySelector('.upload-content').style.display = 'none';
        fileInfo.style.display = 'flex';

        // Store file for form submission
        fileInput.files = new DataTransfer().files;
        const dt = new DataTransfer();
        dt.items.add(file);
        fileInput.files = dt.files;

        // Show success toast with file info
        Swal.fire({
            icon: 'success',
            title: 'File Uploaded!',
            text: `${file.name} (${fileSizeMB} MB) ready for translation`,
            toast: true,
            position: 'top-end',
            showConfirmButton: false,
            timer: 3000,
            timerProgressBar: true,
            iconColor: '#1E40AF',
            customClass: {
                popup: 'colored-toast'
            }
        });
    } else {
        Swal.fire({
            icon: 'error',
            title: 'Invalid File',
            text: 'Please select a valid PowerPoint file (.pptx)',
            confirmButtonColor: '#1E40AF'
        });
    }
}

// Remove file
function removeFile() {
    fileInput.value = '';
    fileName.textContent = '';
    document.querySelector('.upload-content').style.display = 'block';
    fileInfo.style.display = 'none';
}

// Don't show API key input - use server-side keys
serviceSelect.addEventListener('change', (e) => {
    // API keys are handled server-side, no need for user input
    apiKeySection.style.display = 'none';
    apiKeySection.querySelector('input').required = false;
});

// Form submission
const translationForm = document.getElementById('translationForm');
const progressSection = document.getElementById('progressSection');
const successSection = document.getElementById('successSection');
const errorSection = document.getElementById('errorSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const progressDetails = document.getElementById('progressDetails');
const warningMessage = document.getElementById('warningMessage');
const progressLanguages = document.getElementById('progressLanguages');
const successLanguages = document.getElementById('successLanguages');

let currentLanguages = [];

function extractFilename(disposition) {
    if (!disposition) {
        return null;
    }

    const utfMatch = /filename\*?=UTF-8''([^;]+)/i.exec(disposition);
    if (utfMatch && utfMatch[1]) {
        try {
            return decodeURIComponent(utfMatch[1]);
        } catch (error) {
            console.warn('Failed to decode filename from header:', error);
            return utfMatch[1];
        }
    }

    const asciiMatch = /filename="?([^";]+)"?/i.exec(disposition);
    if (asciiMatch && asciiMatch[1]) {
        return asciiMatch[1];
    }

    return null;
}

function renderLanguageBadges(container, languages, state) {
    if (!container) {
        return;
    }

    container.innerHTML = '';
    if (!languages || !languages.length) {
        return;
    }

    const stateLabels = {
        pending: 'In progress',
        success: 'Ready',
        error: 'Failed'
    };

    languages.forEach(lang => {
        const badge = document.createElement('span');
        badge.className = 'language-tag';
        if (state) {
            badge.dataset.state = state;
            if (stateLabels[state]) {
                badge.title = stateLabels[state];
            }
        }
        badge.textContent = lang.toUpperCase();
        container.appendChild(badge);
    });
}

function clearLanguageBadges() {
    if (progressLanguages) {
        progressLanguages.innerHTML = '';
    }
    if (successLanguages) {
        successLanguages.innerHTML = '';
    }
}

translationForm.addEventListener('submit', async (e) => {
    e.preventDefault();

    if (!fileInput.files.length) {
        Swal.fire({
            icon: 'warning',
            title: 'No File Selected',
            text: 'Please select a file to translate',
            confirmButtonColor: '#1E40AF'
        });
        return;
    }

    // Gather language selections
    const selectedLangs = targetLangChoices
        ? targetLangChoices.getValue(true)
        : Array.from(targetLangSelect.selectedOptions).map(option => option.value);
    if (!selectedLangs.length) {
        Swal.fire({
            icon: 'warning',
            title: 'No Language Selected',
            text: 'Please select at least one target language',
            confirmButtonColor: '#1E40AF'
        });
        return;
    }

    // Gather format selections
    const selectedFormats = Array.from(document.querySelectorAll('input[name="formats"]:checked')).map(input => input.value);
    if (!selectedFormats.length) {
        Swal.fire({
            icon: 'warning',
            title: 'No Format Selected',
            text: 'Please select at least one output format',
            confirmButtonColor: '#1E40AF'
        });
        return;
    }

    currentLanguages = [...selectedLangs];
    renderLanguageBadges(progressLanguages, currentLanguages, 'pending');
    renderLanguageBadges(successLanguages, [], null);

    // Hide form and show progress
    translationForm.style.display = 'none';
    progressSection.style.display = 'block';

    if (currentLanguages.length) {
        progressDetails.textContent = `Preparing translations: ${currentLanguages.map(lang => lang.toUpperCase()).join(' · ')}`;
    }

    // Create form data
    const formData = new FormData();
    formData.append('file', fileInput.files[0]);
    selectedLangs.forEach(lang => formData.append('target_langs', lang));
    selectedFormats.forEach(format => formData.append('formats', format));
    formData.append('service', serviceSelect.value);

    // Don't send API key from frontend - use server-side keys
    // API keys are loaded from .env file on the server

    // Declare intervals outside try block
    let uploadInterval;
    let translateInterval;
    let progress = 0;

    try {
        // Initial upload progress (fast)
        progressText.textContent = 'Uploading file...';
        uploadInterval = setInterval(() => {
            if (progress < 20) {
                progress += 5;
                progressFill.style.width = `${progress}%`;
            } else {
                clearInterval(uploadInterval);
                progressText.textContent = 'Processing presentation...';

                // Slower progress for translation
                translateInterval = setInterval(() => {
                    if (progress < 85) {
                        // Slower increments
                        progress += Math.random() * 2 + 0.5;
                        progressFill.style.width = `${Math.min(progress, 85)}%`;

                        if (progress < 40) {
                            progressText.textContent = 'Analyzing slides...';
                            progressDetails.textContent = 'Reading presentation structure';
                        } else if (progress < 60) {
                            progressText.textContent = 'Translating content...';
                            progressDetails.textContent = 'Processing text and tables';
                        } else if (progress < 80) {
                            progressText.textContent = 'Preserving formatting...';
                            progressDetails.textContent = 'Maintaining layout and styles';
                        } else {
                            progressText.textContent = 'Finalizing translation...';
                            progressDetails.textContent = 'Generating output file';
                        }
                    } else {
                        // Hold at 85% until server responds
                        clearInterval(translateInterval);
                    }
                }, 800); // Slower updates
            }
        }, 100);

        // Send request
        const response = await fetch('/translate', {
            method: 'POST',
            body: formData
        });

        // Clear all intervals
        if (uploadInterval) clearInterval(uploadInterval);
        if (translateInterval) clearInterval(translateInterval);

        if (response.ok) {
            const contentType = response.headers.get('Content-Type') || '';
            if (contentType.includes('application/json')) {
                const data = await response.json();
                throw new Error(data.error || 'Translation failed');
            }

            const warnings = response.headers.get('X-Translation-Warnings');
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const disposition = response.headers.get('Content-Disposition');
            const downloadName = extractFilename(disposition) || 'translations.zip';

            // Complete progress smoothly
            let finalProgress = progress;
            const completeInterval = setInterval(() => {
                if (finalProgress < 100) {
                    finalProgress += 5;
                    progressFill.style.width = `${finalProgress}%`;
                    progressText.textContent = 'Translation complete!';
                    progressDetails.textContent = 'Preparing download...';
                } else {
                    clearInterval(completeInterval);

                    // Show success after delay
                    setTimeout(() => {
                        progressSection.style.display = 'none';
                        successSection.style.display = 'block';

                        // Set download link
                        const downloadBtn = document.getElementById('downloadBtn');
                        downloadBtn.href = url;
                        downloadBtn.download = downloadName;

                        renderLanguageBadges(progressLanguages, currentLanguages, 'success');
                        renderLanguageBadges(successLanguages, currentLanguages, 'success');

                        // Show success notification with SweetAlert
                        Swal.fire({
                            icon: 'success',
                            title: 'Translation Complete!',
                            text: `Your presentation has been successfully translated to ${currentLanguages.length} language(s)`,
                            confirmButtonText: 'Download Now',
                            confirmButtonColor: '#1E40AF',
                            showCancelButton: true,
                            cancelButtonText: 'Close',
                            cancelButtonColor: '#6c757d'
                        }).then((result) => {
                            if (result.isConfirmed) {
                                // Trigger download
                                downloadBtn.click();
                            }
                        });

                        if (warningMessage) {
                            if (warnings) {
                                warningMessage.textContent = warnings;
                                warningMessage.style.display = 'block';
                            } else {
                                warningMessage.textContent = '';
                                warningMessage.style.display = 'none';
                            }
                        }
                    }, 500);
                }
            }, 50);
        } else {
            throw new Error('Translation failed');
        }
    } catch (error) {
        // Clear any running intervals
        if (uploadInterval) clearInterval(uploadInterval);
        if (translateInterval) clearInterval(translateInterval);

        // Show error with details
        console.error('Translation error:', error);
        progressSection.style.display = 'none';
        errorSection.style.display = 'block';
        document.getElementById('errorMessage').textContent =
            error.message || 'An error occurred during translation. Please try again.';

        if (warningMessage) {
            warningMessage.textContent = '';
            warningMessage.style.display = 'none';
        }

        renderLanguageBadges(progressLanguages, currentLanguages, 'error');
        renderLanguageBadges(successLanguages, [], null);

        // Show error notification with SweetAlert
        Swal.fire({
            icon: 'error',
            title: 'Translation Failed',
            text: error.message || 'An error occurred during translation. Please try again.',
            confirmButtonColor: '#1E40AF',
            confirmButtonText: 'Try Again'
        });
    }
});

// Reset form
function resetForm() {
    translationForm.style.display = 'block';
    progressSection.style.display = 'none';
    successSection.style.display = 'none';
    errorSection.style.display = 'none';

    removeFile();
    translationForm.reset();
    progressFill.style.width = '0%';
    currentLanguages = [];
    clearLanguageBadges();
    progressText.textContent = 'Initializing...';
    progressDetails.textContent = '';

    if (targetLangChoices) {
        targetLangChoices.removeActiveItems();
        if (defaultLangSelection.length) {
            defaultLangSelection.forEach(value => targetLangChoices.setChoiceByValue(value));
        }
    }

    if (warningMessage) {
        warningMessage.textContent = '';
        warningMessage.style.display = 'none';
    }
}

// Navbar scroll effect
window.addEventListener('scroll', () => {
    const navbar = document.querySelector('.navbar');
    if (window.scrollY > 100) {
        navbar.style.background = 'rgba(255, 255, 255, 0.96)';
        navbar.style.boxShadow = '0 12px 24px -16px rgba(15, 23, 42, 0.3)';
    } else {
        navbar.style.background = 'rgba(255, 255, 255, 0.92)';
        navbar.style.boxShadow = '0 1px 2px 0 rgba(15, 23, 42, 0.08)';
    }
});
