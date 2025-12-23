import React from 'react';
import { Chart } from 'primereact/chart';


interface ChartProps {
    months: string[];
    inProgressData: number[];
    rejectedData: number[];
    completedData: number[];
    returnedData: number[];
    cancelledData: number[];
}

const MyChartComponent: React.FC<ChartProps> = ({
    months,
    inProgressData,
    rejectedData,
    completedData,
    returnedData,
    cancelledData
}) => {
    const chartData = {
        labels: months,
        datasets: [
            { label: "In-Progress", data: inProgressData, backgroundColor: "#42A5F5", fill: false },
            { label: "Rejected", data: rejectedData, backgroundColor: "#EF5350", fill: false },
            { label: "Returned", data: returnedData, backgroundColor: "#FFA726", fill: false },
            { label: "Cancelled", data: cancelledData, backgroundColor: "#B0BEC5", fill: false },
            { label: "Completed", data: completedData, backgroundColor: "#66BB6A", fill: false }
        ]
    };

    const chartOptions = {
        responsive: true,
        // maintainAspectRatio:false,
        plugins: {
            legend: {
                position: 'right',
                labels: {
                    color: 'white',
                    padding: 20,
                    font: {
                        size: 12,
                        weight: 300
                    }
                }
            },
            tooltip: {
                bodyFont: {
                    size: 12
                },
                titleFont: {
                    size: 12
                }
            }
        },
        scales: {
            x: {
                stacked: true,
                ticks: {
                    color: '#ffffff',
                    font: {
                        size: 12
                    }
                },
                grid: {
                    color: '#444444'
                }
            },
            y: {
                stacked: true,
                suggestedMax: 90,
                ticks: {
                    color: '#ffffff',
                    font: {
                        size: 12
                    }
                },
                grid: {
                    color: '#444444'
                }
            }
        }
    };

    return <Chart type="bar" data={chartData} options={chartOptions} style={{ width: '750px', height: '360px', boxShadow: '0 4px 12px rgba(135, 135, 135, 0.25)', border: '1px solid #4A4A4A', padding: '10px 0px 10px 10px' }} />
};

export default MyChartComponent;