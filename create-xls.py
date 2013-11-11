import xlwt
from itertools import chain
import pandas as pd

n_measurements = 3
points_desc_filename = '../points-desc.csv'
points = {}
dists  = {}
rep_dists = {}

# Import points

points_desc = pd.read_csv(points_desc_filename)

points['dorsal'] = points_desc['ponto.dorsal'][points_desc['ponto.dorsal'].notnull()]
points['ventral'] = points_desc['ponto.ventral'][points_desc['ponto.dorsal'].notnull()]
dists['dorsal'] = points_desc['controle.dorsal'].dropna().map(lambda x: x.split('.')).values
dists['ventral'] = points_desc['controle.dorsal'].dropna().map(lambda x: x.split('.')).values
rep_dists['dorsal'] = points_desc['calc.dist.dorsal'].dropna().map(lambda x: x.split('.')).values
rep_dists['ventral'] = points_desc['calc.dist.ventral'].dropna().map(lambda x: x.split('.')).values
final_dists = points_desc['seq.dist.mx'].dropna().map(lambda x: x.split('.')).values

points_cells = {}
dists_cells = {}

header_line = 1
info_col = 1
vistas_col = info_col + 1
n_vistas = len(points.keys())
n_dists = len(dists.keys())
dists_col = vistas_col + 7*n_vistas
rep_col = dists_col + 4*n_dists
final_dists_col = rep_col + 3*n_vistas

# Maximum number of points in a vista
max_n_points = max([len(ps) for ps in chain(points.values(), dists.values())] + [len(final_dists)])
line = -1

new_xls = xlwt.Workbook()
data = new_xls.add_sheet('Dados')
# Write header

def write_header(info, line, data, dists, points):

    data.write(line, 0, info)

    for vi, vista in enumerate(points.keys()):
        data.write(line, vistas_col + 7*vi, vista)
        data.write(line, vistas_col + 7*vi + 1, 'x')
        data.write(line, vistas_col + 7*vi + 2, 'y')
        data.write(line, vistas_col + 7*vi + 3, 'z')
        data.write(line, vistas_col + 7*vi + 4, 'x')
        data.write(line, vistas_col + 7*vi + 5, 'y')
        data.write(line, vistas_col + 7*vi + 6, 'z')

    for vi, vista in enumerate(dists.keys()):
        data.write(line, dists_col + 4*vi, "dist. " + vista)
        data.write(line, dists_col + 4*vi + 1, 'Medida 1')
        data.write(line, dists_col + 4*vi + 2, 'Medida 2')
        data.write(line, dists_col + 4*vi + 3, u'Diferença')

    for vi, vista in enumerate(rep_dists.keys()):
        data.write(line, rep_col + 3*vi, "dist. " + vista)
        data.write(line, rep_col + 3*vi + 1, 'Medida 1')
        data.write(line, rep_col + 3*vi + 2, 'Medida 2')

    data.write(line, final_dists_col, 'Dist. Final')
    data.write(line, final_dists_col + 1, 'Valor')

for mi in range(n_measurements):
    line += mi*max_n_points + 2
    write_header(mi, line - 1, data, dists, points)

    # Create points tables
    for vi, vista in enumerate(points.keys()):
        vi_col = vistas_col + 7*vi
        for pi, p in enumerate(points[vista]):
            # Write point p in (line + pi, vi_col)
            data.write(line + pi, vi_col, p)

    # Create distances tables
    for vi, vista in enumerate(dists.keys()):
        vi_col = dists_col + 4*vi
        for di, d in enumerate(dists[vista]):
            # Write distance d in (line + di, vi_col)
            data.write(line + di, vi_col, '-'.join(d))

    # Create repeatability tables
    for vi, vista in enumerate(rep_dists.keys()):
        vi_col = rep_col + 3*vi
        for di, d in enumerate(rep_dists[vista]):
            # Write
            data.write(line + di, vi_col, '-'.join(d))

    # Create final data table
    for di, d in enumerate(final_dists):
        data.write(line + di, final_dists_col, '-'.join(d))


