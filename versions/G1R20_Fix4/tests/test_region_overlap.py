"""Geometry reference for the split-edge region validation.

Rectangle acceptance is checked against an independent interval intersection
oracle. Orthogonal concave polygons use exact occupied grid cells as an oracle.
Real VBA edge cases are separately executed by the owned Excel input matrix.
"""
from pathlib import Path
import itertools
import unittest

ROOT = Path(__file__).resolve().parents[1]
TOL = 1e-8


def cross(a, b, p):
    return (b[0]-a[0])*(p[1]-a[1])-(b[1]-a[1])*(p[0]-a[0])


def on_segment(p, a, b):
    dx,dy=b[0]-a[0],b[1]-a[1]
    length2=dx*dx+dy*dy
    if length2 <= TOL*TOL: return False
    t=((p[0]-a[0])*dx+(p[1]-a[1])*dy)/length2
    return -TOL <= t <= 1+TOL and abs(cross(a,b,p))/length2**.5 <= TOL


def edges(polygon):
    return list(zip(polygon,polygon[1:]+polygon[:1]))


def on_boundary(p, polygon):
    return any(on_segment(p,a,b) for a,b in edges(polygon))


def inside(p, polygon):
    if on_boundary(p,polygon): return True
    answer=False
    for a,b in edges(polygon):
        if (a[1]>p[1]) != (b[1]>p[1]):
            x=a[0]+(p[1]-a[1])*(b[0]-a[0])/(b[1]-a[1])
            if p[0]<x: answer=not answer
    return answer


def split_at_definitions(polygon, points):
    expanded=[]
    for a,b in edges(polygon):
        values=[p for p in points if on_segment(p,a,b)]
        values.sort(key=lambda p:(p[0]-a[0])*(b[0]-a[0])+(p[1]-a[1])*(b[1]-a[1]))
        expanded.extend(values[:-1])
    return expanded


def area(polygon):
    return abs(sum(a[0]*b[1]-b[0]*a[1] for a,b in edges(polygon)))/2


def accepts(a, b, mids=True):
    points=list(dict.fromkeys(a+b))
    a,b=[split_at_definitions(p,points) for p in [a,b]]
    for (p,q),(r,s) in itertools.combinations(edges(a)+edges(b),2):
        if cross(p,q,r)*cross(p,q,s)<-1e-20 and cross(r,s,p)*cross(r,s,q)<-1e-20:
            return False
    for polygon,other in [(a,b),(b,a)]:
        samples=polygon[:]
        if mids: samples += [((p[0]+q[0])/2,(p[1]+q[1])/2) for p,q in edges(polygon)]
        states={inside(p,other) for p in samples if not on_boundary(p,other)}
        if states=={False,True}: return False
        if not states and abs(area(polygon)-area(other))<TOL: return False
    return True


def rectangle(bounds):
    x0,y0,x1,y1=bounds
    return [(x0,y0),(x1,y0),(x1,y1),(x0,y1)]


def rectangle_oracle(a,b):
    overlap=max(0,min(a[2],b[2])-max(a[0],b[0]))*max(0,min(a[3],b[3])-max(a[1],b[1]))
    if overlap==0: return True
    a_contains=all([a[0]<=b[0],a[1]<=b[1],a[2]>=b[2],a[3]>=b[3]])
    b_contains=all([b[0]<=a[0],b[1]<=a[1],b[2]>=a[2],b[3]>=a[3]])
    if a_contains and b_contains: return False
    return a_contains or b_contains


class RegionOverlapReference(unittest.TestCase):
    def test_rectangles_against_independent_intersection_oracle(self):
        first=(0,0,1,1)
        offsets=[-2,-1,-.999,-.75,-.5,-.01,0,.001,.25,.5,.9,.999,1,1.001,2]
        sizes=[(.25,.5),(.5,1),(1,1),(2,1),(1,2),(2,2)]
        count=old_misses=0
        for x,y,(width,height) in itertools.product(offsets,offsets,sizes):
            second=(x,y,x+width,y+height)
            expected=rectangle_oracle(first,second)
            a,b=rectangle(first),rectangle(second)
            self.assertEqual(accepts(a,b),expected,(first,second))
            old_misses += accepts(a,b,False)!=expected
            count+=1
        self.assertEqual(count,1350)
        self.assertGreater(old_misses,10)

    def test_concave_l_shape_against_exact_cell_occupancy(self):
        # L is the union [0,3]x[0,1] and [0,1]x[0,3]. Each cell has side .5.
        l_shape=[(0,0),(3,0),(3,1),(1,1),(1,3),(0,3)]
        l_cells={(x,y) for x in range(6) for y in range(6) if x<2 or y<2}
        for x0,y0,w,h in itertools.product(range(-2,8),range(-2,8),[1,2,4],[1,2,4]):
            b_cells={(x,y) for x in range(x0,x0+w) for y in range(y0,y0+h)}
            overlap=l_cells & b_cells
            expected=not overlap or (l_cells < b_cells or b_cells < l_cells)
            box=rectangle((x0/2,y0/2,(x0+w)/2,(y0+h)/2))
            self.assertEqual(accepts(l_shape,box),bool(expected),(x0,y0,w,h))

    def test_vba_classifies_expanded_edges_before_region_assignment(self):
        source=(ROOT/'src/vba/RPX_Definition.bas').read_text(encoding='utf8')
        prepare=source.split('Public Sub RPX_PrepareDefinition()',1)[1].split('End Sub',1)[0]
        self.assertLess(prepare.index('expanded.Add points(j)'),prepare.index('Call ValidateRegions'))
        validate=source.split('Private Sub ValidateRegions()',1)[1].split('End Sub',1)[0]
        self.assertIn('nextPoint = DRegions(r).points((i + 1) Mod DRegions(r).pointCount)',validate)
        self.assertIn('If Not OnRegionBoundaryXY(x, y, s) Then',validate)
        self.assertIn('If RPX_InRegion(x, y, s) Then inside = True Else outside = True',validate)
        self.assertEqual(source,(ROOT/'import/RPX_Definition.bas').read_text(encoding='cp932'))


if __name__=='__main__': unittest.main()
