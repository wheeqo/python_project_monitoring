# %%
from project_monitoring import df_late_list, fig_late, fig_prog, sc_project1, sc_project2, sc_project3, sc_project4
from jinja2 import Environment, FileSystemLoader
import plotly

# %%
# Declare Jinja Environment
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("project_report_template.html")
# Define Variables to be Included in HTML Template
template_vars = {"title" : "Project Monitoring - PT. Bina Rekacipta Utama",
                 "fig_prog": plotly.offline.plot(fig_prog, include_plotlyjs=False, output_type='div'),
                 "fig_late": plotly.offline.plot(fig_late, include_plotlyjs=False, output_type='div'),
                 "df_late_list": df_late_list.reset_index(drop=True).to_html(classes='table'),
                 "sc_project1": plotly.offline.plot(sc_project1, include_plotlyjs=False, output_type='div'),
                 "sc_project2": plotly.offline.plot(sc_project2, include_plotlyjs=False, output_type='div'),
                 "sc_project3": plotly.offline.plot(sc_project3, include_plotlyjs=False, output_type='div'),
                 "sc_project4": plotly.offline.plot(sc_project4, include_plotlyjs=False, output_type='div'),
                 }
# Render Template into HTML
html_out = template.render(template_vars)
# Save to .html
with open("project_report.html", "w") as fh:
    fh.write(html_out)

# %%