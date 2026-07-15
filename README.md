# FZERO_BISECT

A native Excel LAMBDA implementation of the bisection root-finding method.

## Purpose

Returns a root of:

```math
f(x)=0
```

within a specified interval using the bisection method.

The interval must bracket a root:

```math
f(lo) \times f(hi)<0
```

## Syntax

```excel
FZERO_BISECT(
    f,
    lo,
    hi,
    Niter
)
```

### Arguments

| Argument | Description |
|-----------|-----------|
| f | Single-variable LAMBDA function |
| lo | Lower bound of search interval |
| hi | Upper bound of search interval |
| Niter | Number of bisection iterations |

## Definition
The BISECTION formula can be added as a named formula with the following content entered as a new object called FZERO_BISECT in the 'Name Manager' located on the 'Formulas" ribbon in the 'defined names' section under 'Name Manager' copy the following in to the "refers to" box
```excel
=LAMBDA(f,lo,hi,Niter,LET(fa,f(lo),fb,f(hi),ITER,LAMBDA(self,k,a,b,fa,fb,LET(mid,(a+b)/2,fm,f(mid),IF(OR(k=0,fm=0),mid,IF(fa*fm < 0,self(self,k-1,a,mid,fa,fm),self(self,k-1,mid,b,fm,fb))))),ITER(ITER,Niter,lo,hi,fa,fb)))
```

OR you can you can define it directly inline when you want to use it
for example, to find the root that lies between 0 and 10 for the equation:

```math
x^2-3=0
```
Using 100 bisections:

```excel
=LET(
  BISECT,
    LAMBDA(f,lo,hi,Niter,LET(fa,f(lo),fb,f(hi),ITER,LAMBDA(self,k,a,b,fa,fb,LET(mid,(a+b)/2,fm,f(mid),IF(OR(k=0,fm=0),mid,IF(fa*fm<0,self(self,k-1,a,mid,fa,fm),self(self,k-1,mid,b,fm,fb))))),ITER(ITER,Niter,lo,hi,fa,fb))),
  BISECT(LAMBDA(x,x^2-3),0,10,100)
  )
```
I recommend adding it to the name space as it's quite large


## Example 1


```math
x^2-2=0
```

```excel
=FZERO_BISECT(
  LAMBDA(x,
    x^2-2
  ),
  1,
  2,
  100)
```

Result:

```text
1.414213562373095
```

## Example 2

Find:

```math
\cos(x)-x=0
```

```excel
=FZERO_BISECT(
    LAMBDA(x,COS(x)-x),
    0,
    1,
    100
)
```

Result:

```text
0.739085133215161
```

## Example 3

Crank Slab Eigenvalue

Solve:

```math
\beta\tan(\beta)-Bi=0
```

```excel
=FZERO_BISECT(
    LAMBDA(x,x*TAN(x)-2),
    0.01,
    PI()/2,
    100
)
```

## Notes

- Implemented entirely using native Excel LAMBDA functions.
- Requires Microsoft Excel 365.
- No VBA required.
- No Solver required.
- Convergence is guaranteed when the root is bracketed.
- Particularly useful for engineering and scientific calculations involving transcendental equations.

## Author

Richard Edmonds
