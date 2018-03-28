Function LAMTHETA(x, r)

  Dim k As Integer
  LAMTHETA = 1
  For k = 1 To r
      LAMTHETA = LAMTHETA + (x ^ k * (r - (k - 1)) ^ k) / (Application.Fact(k))
  Next k

End Function


Function LAMBERTW(x, r As Integer, n As Integer)

  If n = 1 Then
      LAMBERTW = 1 / r * Application.Ln(LAMTHETA(x, r))
  Else
      LAMBERTW = 1 / r * Application.Ln((LAMBERTW(x, r, n - 1) * (1 + LAMBERTW(x, r, n - 1))) / x * LAMTHETA(x, r))
  End If

End Function
