
// Store original data
const originalData = {treemap_data_json};

var ctx = document.getElementById('myChart').getContext('2d');
var myChart = new Chart(ctx, {{
    type: 'treemap',
    data: {{
        datasets: [
            {{
                label: 'Afinidades',
                tree: originalData,
                key: 'value',
                groups: ['name'],
                backgroundColor: (ctx) => {{
                    const HUE_START = 200; // Blue
                    const HUE_END = 0; // Red
                    const value = ctx.raw.v;
                    const max = ctx.dataset.tree.reduce((max, node) => Math.max(max, node.value), -Infinity);
                    const hue = HUE_START + (HUE_END - HUE_START) * (value / max);
                    return `hsl(${{hue}}, 70%, 60%)`;
                }},
                spacing: 1,
                borderWidth: 2,
                borderColor: 'white'
            }}
        ]
    }},
    options: {{
        maintainAspectRatio: false,
        plugins: {{
            title: {{
                display: true,
                text: 'Afinidades por Agente'
            }},
            legend: {{
                display: false
            }},
            tooltip: {{
                callbacks: {{
                    label: function(context) {{
                        const item = context.raw;
                        return `${{item.g}}: ${{item.v}}`;
                    }}
                }}
            }}
        }}
    }}
}});

// Search functionality
const searchInput = document.getElementById('searchInput');
searchInput.addEventListener('keyup', (e) => {{
    const searchValue = e.target.value.toLowerCase();
    const filteredData = originalData.filter(item =>
        item.name.toLowerCase().includes(searchValue)
    );
    myChart.data.datasets[0].tree = filteredData;
    myChart.update();
}});
