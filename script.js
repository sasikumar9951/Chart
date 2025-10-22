document
  .getElementById("excel-file-input")
  .addEventListener("change", handleFile);

let myChart;

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      // MODIFIED: Ensure dates are read correctly if Excel has date cells
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { cellDates: true });

      if (jsonData.length > 0) {
        const chartData = processParticipationData(jsonData);
        createOrUpdateChart(chartData);
        document.getElementById("download-btn").style.display = "block";
      } else {
        alert("Excel sheet kaaliyaga ulladhu!");
      }
    } catch (error) {
      console.error("File process pannum podhu thavaru:", error);
      alert("Sariyana Excel format-il file ullatha ena sari paarkavum.");
    }
  };
  reader.readAsArrayBuffer(file);
}

// MODIFIED: Intha function-a mothama maathirukkom
function processParticipationData(jsonData) {
  const participationData = {};
  const allSections = new Set();
  
  // NEW: Month names array, idhu labels create panna use aagum
  const monthNames = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];

  jsonData.forEach((row) => {
    const year = row.Year;
    // NEW: Month data-va Excel-lendhu edukkrom (Number 1-12 nu assume panrom)
    const month = row.Month; 
    const section = row.Section;
    const status = row.Status;

    // NEW: Month data-vum irukka-nu check panrom
    if (status && status.toLowerCase() === "participated" && year && month && section) {
      
      // NEW: Month number sariya irukka-nu check panrom (1-12)
      if (month < 1 || month > 12) {
        console.warn("Invalid month number:", month, row);
        return; // Thavarana month-a skip panrom
      }

      // NEW: "2024-Jan", "2024-Feb" mathiri label create panrom
      const yearMonthLabel = `${year}-${monthNames[month - 1]}`;

      allSections.add(section);
      
      // NEW: Data-va year[section] ku pathila yearMonth[section]-la store panrom
      if (!participationData[yearMonthLabel]) {
        participationData[yearMonthLabel] = {};
      }
      if (!participationData[yearMonthLabel][section]) {
        participationData[yearMonthLabel][section] = 0;
      }
      participationData[yearMonthLabel][section]++;
    }
  });

  // NEW: Labels ippo '2023-Jan', '2023-Feb' mathiri irukkum
  const labels = Object.keys(participationData);
  const sections = Array.from(allSections).sort();

  // NEW: Labels-a alphabetical-a sort panrathukku pathila, chronological-a (kaala varisaippadi) sort panrom
  labels.sort((a, b) => {
    const [yearA, monthStrA] = a.split('-');
    const [yearB, monthStrB] = b.split('-');
    
    const monthA = monthNames.indexOf(monthStrA);
    const monthB = monthNames.indexOf(monthStrB);

    if (yearA !== yearB) {
      return parseInt(yearA) - parseInt(yearB); // Year vechu sort panrom
    }
    return monthA - monthB; // Same year-a irundha, month vechu sort panrom
  });


  const colors = [
    "rgba(255, 99, 132, 0.7)",
    "rgba(54, 162, 235, 0.7)",
    "rgba(255, 206, 86, 0.7)",
    "rgba(75, 192, 192, 0.7)",
    "rgba(153, 102, 255, 0.7)",
    "rgba(255, 159, 64, 0.7)",
  ];

  const datasets = sections.map((section, index) => {
    // MODIFIED: 'year'-ku pathila 'yearMonthLabel' use panrom
    const data = labels.map((yearMonthLabel) => {
      return participationData[yearMonthLabel][section] || 0;
    });
    return {
      label: `Section ${section}`,
      data: data,
      backgroundColor: colors[index % colors.length],
      borderColor: colors[index % colors.length].replace("0.7", "1"),
      borderWidth: 1,
    };
  });

  return { labels, datasets };
}

function createOrUpdateChart(chartData) {
  const ctx = document.getElementById("myChart").getContext("2d");

  if (myChart) {
    myChart.destroy();
  }

  myChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: chartData.labels,
      datasets: chartData.datasets,
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          title: {
            display: true,
            text: "Number of Students Participated",
          },
        },
        x: {
          title: {
            display: true,
            // MODIFIED: X-axis title-a maathirukkom
            text: "Year-Month",
          },
        },
      },
      plugins: {
        title: {
          display: true,
          // MODIFIED: Title-a konjam update pannirukkom
          text: "CSE Department Event Participation (Month Wise)",
          font: {
            size: 18,
          },
        },
        legend: {
          display: true,
          position: "top",
        },
        tooltip: {
          mode: "index",
          intersect: false,
        },
      },
    },
  });
}

document.getElementById("download-btn").addEventListener("click", function () {
  if (myChart) {
    const canvas = myChart.canvas;
    const ctx = canvas.getContext("2d");
    ctx.globalCompositeOperation = "destination-over";
    ctx.fillStyle = "white";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    const link = document.createElement("a");
    link.href = canvas.toDataURL("image/jpeg", 0.9);
    link.download = "student-participation-chart.jpg";
    link.click();
    ctx.globalCompositeOperation = "source-over";
  }
});
