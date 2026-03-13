let data = [];

// Load Excel automatically when page starts
fetch("data.xlsx")
.then(res => res.arrayBuffer())
.then(buffer => {

    const workbook = XLSX.read(buffer, {type:"array"});

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // Convert Excel → JSON automatically
    data = XLSX.utils.sheet_to_json(sheet);

    initialize();
});


// Initialize dropdowns
function initialize(){

    const seriesSelect = document.getElementById("series");

    seriesSelect.innerHTML =
        '<option value="">Select Series</option>';

    const seriesList =
        [...new Set(data.map(d => d.Series))];

    seriesList.forEach(s=>{
        let opt=document.createElement("option");
        opt.value=s;
        opt.text=s;
        seriesSelect.add(opt);
    });

    seriesSelect.onchange = updateSizes;
}


// Update Size dropdown
function updateSizes(){

    const selectedSeries =
        document.getElementById("series").value;

    const sizeSelect =
        document.getElementById("size");

    sizeSelect.innerHTML =
        '<option value="">Select Size</option>';

    const sizes = data
        .filter(d => d.Series === selectedSeries)
        .map(d => d.Size);

    sizes.forEach(size=>{
        let opt=document.createElement("option");
        opt.value=size;
        opt.text=size;
        sizeSelect.add(opt);
    });
}


// Lookup values
function findValues(){

    const s =
        document.getElementById("series").value;

    const size =
        document.getElementById("size").value;

    const row = data.find(
        d => d.Series === s && d.Size == size
    );

    if(row){
        pitch.innerText = row.Pitch;
        evalue.innerText = row["Min E-Value"];
    }
    else{
        pitch.innerText="Not found";
        evalue.innerText="Not found";
    }
}
