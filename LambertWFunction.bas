Function LAMPHI(x, r)

  Dim k As Integer
  LAMPHI = 1
  For k = 1 To r
     LAMPHI = LAMPHI + (x ^ k * (r - (k - 1)) ^ k) / (Application.Fact(k))
  Next k

End Function


Function LAMBERTW(x, r As Integer, n As Integer)

  If n = 1 Then
    LAMBERTW = 1 / r * Application.Ln(LAMPHI(x, r))
  Else
    LAMBERTW = 1 / r * Application.Ln((LAMBERTW(x, r, n - 1) * (1 + LAMBERTW(x, r, n - 1))) / x * LAMPHI(x, r))
  End If

End Function
