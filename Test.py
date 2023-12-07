import plotly.express as px

# Example data
labels = ['Category1', 'Category2', 'Category3']
sizes = [30, 40, 30]

# Create an interactive pie chart with hover functionality
fig = px.pie(names=labels, values=sizes, title='Pie Chart with Hover', hover_data=['values'])

# Show the plot
fig.show()
