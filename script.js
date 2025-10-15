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
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

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

function processParticipationData(jsonData) {
  const participationData = {};
  const allSections = new Set();

  jsonData.forEach((row) => {
    const year = row.Year;
    const section = row.Section;
    const status = row.Status;

    if (status && status.toLowerCase() === "participated" && year && section) {
      allSections.add(section);
      if (!participationData[year]) {
        participationData[year] = {};
      }
      if (!participationData[year][section]) {
        participationData[year][section] = 0;
      }
      participationData[year][section]++;
    }
  });

  const labels = Object.keys(participationData).sort();
  const sections = Array.from(allSections).sort();

  const colors = [
    "rgba(255, 99, 132, 0.7)",
    "rgba(54, 162, 235, 0.7)",
    "rgba(255, 206, 86, 0.7)",
    "rgba(75, 192, 192, 0.7)",
    "rgba(153, 102, 255, 0.7)",
    "rgba(255, 159, 64, 0.7)",
  ];

  const datasets = sections.map((section, index) => {
    const data = labels.map((year) => {
      return participationData[year][section] || 0;
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
            text: "Year",
          },
        },
      },
      plugins: {
        title: {
          display: true,
          text: "CSE Department Event Participation (Year & Section Wise)",
          font: {
            size: 18,
          },
        },
        legend: {
          display: true, // Legend ippo thevai, adhanaala 'true' nu maathrom
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
