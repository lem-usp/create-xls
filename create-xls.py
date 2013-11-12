# -*- coding: utf-8 -*-

import xlwt
from itertools import chain
import pandas as pd

n_measurements = 21
points_desc_filename = 'points-desc.csv'
points = {}
dists  = {}
rep_dists = {}

# Import points

points_desc = pd.read_csv(points_desc_filename, sep=';', keep_default_na=False)

points['dorsal'] = points_desc['ponto.dorsal'][points_desc['ponto.dorsal'] != '']
points['ventral'] = points_desc['ponto.ventral'][points_desc['ponto.ventral'] != '']
dists['dorsal'] = points_desc['controle.dorsal'][points_desc['controle.dorsal'] != ''].map(lambda x: x.split('.')).values
dists['ventral'] = points_desc['controle.ventral'][points_desc['controle.ventral'] != ''].map(lambda x: x.split('.')).values
rep_dists['dorsal'] = points_desc['calc.dist.dorsal'][points_desc['calc.dist.dorsal'] != ''].map(lambda x: x.split('.')).values
rep_dists['ventral'] = points_desc['calc.dist.ventral'][points_desc['calc.dist.ventral'] != ''].map(lambda x: x.split('.')).values
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

def get_column_letter(col_idx):
    """Convert a column number into a column letter (3 -> 'C')

    Right shift the column col_idx by 26 to find column letters in reverse
    order.  These numbers are 1-based, and can be converted to ASCII
    ordinals by adding 64.

    """
    # these indicies corrospond to A -> ZZZ and include all allowed
    # columns
    if not 1 <= col_idx <= 18278:
        msg = 'Column index out of bounds: %s' % col_idx
        raise ColumnStringIndexException(msg)
    ordinals = []
    temp = col_idx
    while temp:
        quotient, remainder = divmod(temp, 26)
        # check for exact division and borrow if needed
        if remainder == 0:
            quotient -= 1
            remainder = 26
        ordinals.append(remainder + 64)
        temp = quotient
    ordinals.reverse()
    return ''.join([chr(ordinal) for ordinal in ordinals])

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

form_string = 'IF(AND({}<>0,{}<>0),SQRT(({}-{})^2+({}-{})^2+({}-{})^2),"")'
diff_form_string = "ABS({}-{})"

for mi in range(n_measurements):
    points_dict = {}
    line += mi*max_n_points + 2
    write_header(mi, line - 1, data, dists, points)

    # Create points tables
    for vi, vista in enumerate(points.keys()):
        vi_col = vistas_col + 7*vi
        for pi, p in enumerate(points[vista]):
            points_dict[p] = (line + pi, vi_col)
            # Write point p in (line + pi, vi_col)
            data.write(line + pi, vi_col, p)

    # Create distances tables
    for vi, vista in enumerate(dists.keys()):
        vi_col = dists_col + 4*vi
        for di, d in enumerate(dists[vista]):
            print d
            # Write distance d in (line + di, vi_col)
            data.write(line + di, vi_col, '-'.join(d))
            # Write formula for distance
            p1_x = get_column_letter(vistas_col + 7*vi + 2) + str(points_dict[d[0]][0] + 1)
            p1_y = get_column_letter(vistas_col + 7*vi + 3) + str(points_dict[d[0]][0] + 1)
            p1_z = get_column_letter(vistas_col + 7*vi + 4) + str(points_dict[d[0]][0] + 1)

            p2_x = get_column_letter(vistas_col + 7*vi + 2) + str(points_dict[d[1]][0] + 1)
            p2_y = get_column_letter(vistas_col + 7*vi + 3) + str(points_dict[d[1]][0] + 1)
            p2_z = get_column_letter(vistas_col + 7*vi + 4) + str(points_dict[d[1]][0] + 1)

            # Measure 1 formula
            fstring = form_string.format(p1_x, p2_x, p1_x, p2_x, p1_y, p2_y, p1_z, p2_z)
            dist_formula = xlwt.Formula(fstring)
            data.write(line + di, vi_col + 1, dist_formula)

            # Measure 2 formula
            p1_x = get_column_letter(vistas_col + 7*vi + 5) + str(points_dict[d[0]][0] + 1)
            p1_y = get_column_letter(vistas_col + 7*vi  + 6) + str(points_dict[d[0]][0] + 1)
            p1_z = get_column_letter(vistas_col + 7*vi  + 7) + str(points_dict[d[0]][0] + 1)

            p2_x = get_column_letter(vistas_col + 7*vi  + 5) + str(points_dict[d[1]][0] + 1)
            p2_y = get_column_letter(vistas_col + 7*vi  + 6) + str(points_dict[d[1]][0] + 1)
            p2_z = get_column_letter(vistas_col + 7*vi  + 7) + str(points_dict[d[1]][0] + 1)

            fstring = form_string.format(p1_x, p2_x, p1_x, p2_x, p1_y, p2_y, p1_z, p2_z)
            dist_formula = xlwt.Formula(fstring)
            data.write(line + di, vi_col + 2, dist_formula)

            # Write difference formula
            diff_1 = get_column_letter(vi_col + 2) + str(line + di + 1)
            diff_2 = get_column_letter(vi_col + 3) + str(line + di + 1)
            diff_formula = xlwt.Formula(diff_form_string.format(diff_1, diff_2))
            data.write(line + di, vi_col + 3, diff_formula)

    # Create repeatability tables
    for vi, vista in enumerate(rep_dists.keys()):
        vi_col = rep_col + 3*vi
        for di, d in enumerate(rep_dists[vista]):
            # Write
            data.write(line + di, vi_col, '-'.join(d))

    # Create final data table
    for di, d in enumerate(final_dists):
        data.write(line + di, final_dists_col, '-'.join(d))

new_xls.save('new.xls')
