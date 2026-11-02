public class Barang {
    private String nama;
    private int harga;
    private int stok;

    public Barang() {
    }

    public Barang(String nama, int harga, int stok) {
        this.nama = nama;
        this.harga = harga;
        this.stok = stok;
    }

    public String getNama() {
        return nama;
    }

    public void setNama(String nama) {
        this.nama = nama;
    }

    public int getHarga() {
        return harga;
    }

    public void setHarga(int harga) {
        this.harga = harga;
    }

    public int getStok() {
        return stok;
    }

    public void setStok(int stok) {
        this.stok = stok;
    }

    public int hitungTotalBayar(int qty) {
        return qty * this.harga;
    }

    public boolean kurangiStok(int qty) {
        if (qty <= this.stok) {
            this.stok -= qty;
            return true;
        }
        return false;
    }
}
