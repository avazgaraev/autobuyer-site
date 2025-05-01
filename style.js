const auctionFiles = {
    'Copart': './locations - copart 2024-11-12 14_00_30.xlsx',
    'IAAI': './location - iaai 2024-11-12 14_00_34.xlsx',
  };
  
  let locationData = [];

  function populateAuctionTypes() {
    const auctionSelect = document.getElementById('auctionSelect');
    auctionSelect.innerHTML = '<option selected disabled>Select Auction Type</option>';
  
    Object.keys(auctionFiles).forEach(type => {
      const option = document.createElement('option');
      option.value = type;
      option.textContent = type;
      auctionSelect.appendChild(option);
    });
  }
  
  function loadLocationsForAuction(auctionType) {
    const filePath = auctionFiles[auctionType];
    if (!filePath) return;
  
    fetch(filePath)
      .then(response => response.arrayBuffer())
      .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
        locationData = [];
  
        const locationSelect = document.getElementById('locationSelect');
        locationSelect.innerHTML = '<option selected disabled>Select Location</option>';
  
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          const fullName = row[1]; // Column E: Name
          const price = row[3];    // Column D: Price
  
          if (fullName && price != null && price !== '') {
            locationData.push({ name: fullName, price });
  
            const option = document.createElement('option');
            option.value = fullName;
            option.textContent = fullName; // ✅ Only Name shown
            locationSelect.appendChild(option);
          }
        }
      })
      .catch(error => {
        console.error(`Error loading ${auctionType} file:`, error);
      });
  }
  
  document.getElementById('auctionSelect').addEventListener('change', function (e) {
    const auctionType = e.target.value;
    loadLocationsForAuction(auctionType);
    document.getElementById('priceDisplay').textContent = 'Qiyməti görmək üçün seçim edin';
  });
  
  document.getElementById('locationSelect').addEventListener('change', function (e) {
    const selectedName = e.target.value;
    const priceEl = document.getElementById('priceDisplay');
  
    const match = locationData.find(item => item.name === selectedName);
    if (match) {
      priceEl.textContent = `Qiymət: $${match.price}`;
    } else {
      priceEl.textContent = '';
    }
  });
  
  window.onload = populateAuctionTypes;