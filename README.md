# TCU - Tree Command for UNO

Output the API tee from the UNO object or IDL name.

## Installing TCU extension

Add <a href="https://github.com/p--q/TCU/tree/master/TCU/oxt">TCU.oxt</a> with Extension Manager.

## Usage

	tcu = smgr.createInstanceWithContext("pq.Tcu", ctx)  # 
	pycomp.wtree(ctx)
	print("\n".join(pycomp.tree(ctx)))

