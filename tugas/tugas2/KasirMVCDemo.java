import java.util.Scanner;

public class KasirMVCDemo {
    public static void main(String[] args) {
        Scanner in = new Scanner(System.in);

        // 1. Inisialisasi Model, View, dan Controller
        Barang model = retriveBarangFromDatabase();
        KasirView view = new KasirView();
        KasirController controller = new KasirController(model, view);

        // 2. Tampilan info barang awal melalui controller
        controller.updateView();

        // 3. Transaksi pembelian
        view.printTransaksiHeader();
        System.out.print("Jumlah barang : ");
        int qty = in.nextInt();

        // 4. Proses transaksi melalui controller dan cetak pembayaran
        controller.prosesTransaksi(qty);

        // 5. Tampilan info barang setelah transaksi diperbarui
        controller.updateView();
    }

    private static Barang retriveBarangFromDatabase() {
        Barang barang = new Barang();
        barang.setNama("GeForce GTX 1060Ti");
        barang.setHarga(3000);
        barang.setStok(10);
        return barang;
    }
}
