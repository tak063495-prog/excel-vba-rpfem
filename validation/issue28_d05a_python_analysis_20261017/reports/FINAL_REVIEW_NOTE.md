# Publication review correction

The independent final review found that the draft README used an output directory outside the bundle for a fresh six-route experiment. The recorded producer intentionally records result paths relative to its bundle root, so that example would fail at `folder.relative_to(ROOT)` after the first child execution.

The published README now documents that existing constraint and uses `results/fixed_point_new` inside the bundle. The producer, numerical sources, thresholds, manifests, and recorded experiment results are unchanged. The saved-vector verification output may still be outside the bundle.

The initial unpublished ZIP, its checksums, and its successful extracted verification records remain preserved outside the repository in the review workspace. A new distribution ZIP was built after this documentation correction; its actual SHA256 is in PACKAGE.json and SHA256SUMS.txt. No additional optimizer calls were made for this documentation change.
