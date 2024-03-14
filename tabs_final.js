
 // Variable to store the current tab index
    let currentTab = 0;

// Function to handle tab navigation and button visibility
function showTab(tabIndex) {
    const nextButton = document.querySelector('.next');
    const preButton = document.querySelector('.pre');
    const tabButtons = document.querySelectorAll('.tab-button');

    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.style.display = 'none';  // Hide all tabs
    });

    document.querySelectorAll('.tab-button').forEach(button => {
        button.classList.remove('active');  // Remove the 'active' class from all tab buttons
    });

    // Hide 'pre' button on tab 1 
    if (tabIndex === 0) {   
        preButton.style.visibility = 'hidden';    
    } else {  
        preButton.style.visibility = 'visible'; 
    }

    // Hide 'next' button on tab 5
    if (tabIndex === 4) { 
        nextButton.style.visibility = 'hidden';   
    } else {  
        nextButton.style.visibility = 'visible';  
    }

    const currentTabId = `tab${tabIndex + 1}`;
    const currentTabContent = document.getElementById(currentTabId);

    currentTabContent.style.display = 'block';  // Show the selected tab
    document.getElementById(`tabButton${tabIndex + 1}`).classList.add('active');  // Add the 'active' class to the clicked tab button

    currentTab = tabIndex;

    // Check for activity link in current tab
    const activityLink = currentTabContent.querySelector('a[target="_blank"]');
    if (activityLink) {
        // Disable navigation to next tabs if activity link exists
        nextButton.style.visibility = 'hidden'; 
        
        for (let i = tabIndex + 1; i < tabButtons.length; i++) {
            tabButtons[i].style.pointerEvents = 'none'; // Disable clicking on next tabs
            tabButtons[i].classList.add('disabled'); // Optionally, add a disabled style to indicate it's disabled
        }
    } else {
        // Enable navigation to all tabs if no activity link exists
        nextButton.style.visibility = 'visible'; 
        for (let i = 0; i < tabButtons.length; i++) {
            tabButtons[i].style.pointerEvents = 'auto'; // Enable clicking on all tabs
            tabButtons[i].classList.remove('disabled'); // Remove disabled style
        }
    }
}











    // Function to handle showing next tab
    function showNextTab() {
        const nextButton = document.querySelector('.next');
        if (currentTab < 4 && nextButton.style.visibility !== 'hidden' ) {
            showTab(currentTab + 1);
        }
    }

    // Function to handle showing previous tab
    function showPreviousTab() {
        const preButton = document.querySelector('.pre');
        if (currentTab > 0 && preButton.style.visibility !== 'hidden') {
            showTab(currentTab - 1);
        }
    }

    // Event listener for external links
    document.querySelectorAll('a[target="_blank"]').forEach(link => {
        link.addEventListener('click', function() {
            const nextButton = document.querySelector('.next');
            nextButton.style.visibility = 'visible'; // Make 'next' button visible again when external link is clicked
        });
    });

    showTab(0); // Show the initial tab when loaded
