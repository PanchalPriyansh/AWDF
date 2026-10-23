import { useEffect, useRef } from 'react';
import Chart from 'chart.js/auto';

/**
 * Heavy Third-Party Analytics Component (Practical 8 - Supplementary Problem 1)
 * Bundles Chart.js (~200KB) into an isolated chunk that is ONLY fetched on-demand
 * when the user explicitly expands the Task Analytics view.
 */
function TaskAnalyticsChart({ tasks = [] }) {
  const chartRef = useRef(null);
  const chartInstance = useRef(null);

  useEffect(() => {
    if (!chartRef.current) return;

    // Destroy prior chart instance if it exists
    if (chartInstance.current) {
      chartInstance.current.destroy();
    }

    const completedCount = tasks.filter((t) => t.completed).length;
    const pendingCount = tasks.length - completedCount;

    const lowCount = tasks.filter((t) => (t.priority || 'medium') === 'low').length;
    const medCount = tasks.filter((t) => (t.priority || 'medium') === 'medium').length;
    const highCount = tasks.filter((t) => (t.priority || 'medium') === 'high').length;

    const ctx = chartRef.current.getContext('2d');
    chartInstance.current = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['Completed', 'Pending', 'High Priority', 'Medium Priority', 'Low Priority'],
        datasets: [
          {
            label: 'Task Distribution',
            data: [completedCount, pendingCount, highCount, medCount, lowCount],
            backgroundColor: [
              '#22c55e', // Green for completed
              '#f59e0b', // Amber for pending
              '#ef4444', // Red for high
              '#3b82f6', // Blue for medium
              '#06b6d4', // Cyan for low
            ],
            borderWidth: 2,
            borderColor: '#ffffff',
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              boxWidth: 14,
              padding: 12,
              font: {
                size: 12,
                family: 'inherit',
              },
            },
          },
          tooltip: {
            callbacks: {
              label: (context) => ` ${context.label}: ${context.raw} tasks`,
            },
          },
        },
      },
    });

    return () => {
      if (chartInstance.current) {
        chartInstance.current.destroy();
        chartInstance.current = null;
      }
    };
  }, [tasks]);

  return (
    <div
      className="task-analytics-card"
      style={{
        backgroundColor: 'var(--card-bg, #ffffff)',
        border: '1px solid var(--border-color, #e2e8f0)',
        borderRadius: '10px',
        padding: '1.5rem',
        marginTop: '1.5rem',
        marginBottom: '1.5rem',
        boxShadow: '0 2px 10px rgba(0,0,0,0.05)',
      }}
    >
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
        <div>
          <h3 style={{ margin: 0, fontSize: '1.15rem', color: '#1d4ed8' }}>
            📊 Live Task Analytics (Lazy-Loaded Chart.js)
          </h3>
          <p style={{ margin: '4px 0 0', fontSize: '0.85rem', color: 'var(--text-muted, #64748b)' }}>
            Rendered using an isolated on-demand Chart.js bundle chunk (Supplementary 1)
          </p>
        </div>
        <span style={{
          fontSize: '0.75rem',
          backgroundColor: '#eff6ff',
          color: '#1d4ed8',
          border: '1px solid #bfdbfe',
          padding: '4px 10px',
          borderRadius: '16px',
          fontWeight: 600,
        }}>
          Dynamic Chunk Loaded
        </span>
      </div>

      <div style={{ position: 'relative', height: '240px', width: '100%' }}>
        <canvas ref={chartRef} />
      </div>
    </div>
  );
}

export default TaskAnalyticsChart;
