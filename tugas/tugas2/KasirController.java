public class KasirController {
    private Barang model;
    private KasirView view;

    public KasirController(Barang model, KasirView view) {
        this.model = model;
        this.view = view;
    }

    public void setBarangNama(String nama) {
        model.setNama(nama);
    }

    public String getBarangNama() {
        return model.getNama();
    }

    public void setBarangHarga(int harga) {
        model.setHarga(harga);
    }

    public int getBarangHarga() {
        return model.getHarga();
    }

    public void setBarangStok(int stok) {
        model.setStok(stok);
    }

    public int getBarangStok() {
        return model.getStok();
    }

    public void updateView() {
        view.printBarangDetails(model.getNama(), model.getHarga(), model.getStok());
    }

    public boolean prosesTransaksi(int qty) {
        if (qty <= 0) {
            view.printPesan("Jumlah pembelian tidak valid!");
            return false;
        }
        if (model.kurangiStok(qty)) {
            int totalBayar = model.hitungTotalBayar(qty);
            view.printPembayaran(totalBayar);
            return true;
        } else {
            view.printPesan("Transaksi Gagal: Stok tidak mencukupi!");
            return false;
        }
    }
}
