from native_common import *
base=saved_mesh();write_input(base,R/'results/base_input.bin')
z=np.load(R/'results/staged_1000_selection.npz');score=z['score'];selected=z['selected']
(R/'results/scores.bin').write_bytes(score.astype('<f8').tobytes())
work=R/'tests'/f'Geometry_{time.time_ns()}.xlsm';shutil.copy2(source,work)
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;book=None
try:
    book=app.Workbooks.Open(str(work),UpdateLinks=0);app.EnableEvents=False;inject(book)
    setvalue(book,'ADAPT_MESH_POLICY','LEGACY_THEN_REFINE');setvalue(book,'ADAPT_REFINE_CAP',1000);setvalue(book,'FS_BRACKET_POLICY','AUDITED_UPPER_LIMIT')
    compile_book(app,book)
    out=app.Run(f"'{work.name}'!MeshRefine",str(R/'results/base_input.bin'),str(R/'results/scores.bin'),str(R/'results/native1000.bin'),1000,.3,True,False)
    assert out.startswith('PASS'),out
    xy,tri,mat=read_geom(R/'results/native1000.bin');expected=refine(base,set(map(int,selected)),True)
    assert canonical(xy,tri,mat)==canonical(expected.xy,expected.tri,expected.material_id)
    m=geometry_mesh(xy,tri,mat);save_mesh(R/'results/native1000_mesh.npz',m);write_input(m,R/'results/native1000_input.bin')
    # Legacy phase must honor explicit override even though the global policy enables red refinement.
    oldscore=np.load(Path('C:/Users/link_/Desktop/RPFEM_G1R10_Mesh_20261007/results/kinematic_scores.npz'))['neighbour']
    (R/'results/old_scores.bin').write_bytes(oldscore.astype('<f8').tobytes())
    legacy=app.Run(f"'{work.name}'!MeshRefine",str(R/'results/base_input.bin'),str(R/'results/old_scores.bin'),str(R/'results/forced_legacy.bin'),5000,60/781,False,True)
    assert legacy.startswith('PASS'),legacy
    assert (R/'results/forced_legacy.bin').read_bytes()==Path('C:/Users/link_/Desktop/RPFEM_G1R10_Mesh_20261007/results/native_original_legacy.bin').read_bytes()
    rec={'native_compile':'PASS','native_refinement':out,'legacy_phase':legacy,'legacy_geometry_matches_G1R9_bytes':True,'native_python_exact_coordinate_triangles_materials':'PASS','native_runtime_mesh_order_used_for_fixed_solves':True}
    (R/'results/native_geometry.json').write_text(json.dumps(rec,indent=2),encoding='utf-8');print(rec,flush=True)
finally:
    if book is not None:book.Close(False)
    app.Quit()
