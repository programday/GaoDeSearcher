function Clip(t, a, e) {
    return Math.min(Math.max(t, a), e)
}

function get_data(m, d, p) {
    var t = 256
      , a = -85.0511287798
      , e = 85.0511287798
      , i = -180
      , n = 180
      , r = 6378137
      , o = Math.PI
      , s = {};
    var f = 2 * o * r;
    m = Clip(m, a, e) * o / 180,
    d = Clip(d, i, n) * o / 180,
    sinLatitude = Math.sin(m);
    var c = r * d
      , l = Math.log((1 + sinLatitude) / (1 - sinLatitude))
      , v = parseInt(r / 2) * l
      , u = t << p
      , h = f / u
      , g = Clip((f / 2 + c) / h + .5, 0, u - 1)
      , y = f / 2 - v;
    y = parseInt(y);
    var _ = Clip(y / h + .5, 0, u - 1);
    return {
        x: parseInt(g),
        y: parseInt(_)
    }
}