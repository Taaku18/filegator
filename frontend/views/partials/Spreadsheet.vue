<template>
  <div>
    <div class="modal-card">
      <header class="modal-card-head">
        <p class="modal-card-title">
          {{ currentItem.name }}
        </p>
      </header>
      <section class="modal-card-body preview">
        <div id="out-table"></div>
      </section>
      <footer class="modal-card-foot">
        <button class="button" type="button" @click="$parent.close()">
          {{ lang('Close') }}
        </button>
      </footer>
    </div>
  </div>
</template>

<script>
    import api from '../../api/api'
    import XLSX from 'xlsx'
    import Spreadsheet from 'x-data-spreadsheet'
    import zhCN from 'x-data-spreadsheet/dist/locale/zh-cn'

    Spreadsheet.locale('zh-cn', zhCN)

    function stox(wb) {
        var out = []
        wb.SheetNames.forEach(function(name) {
            var o = {name:name, rows:{}}
            var ws = wb.Sheets[name]
            var aoa = XLSX.utils.sheet_to_json(ws, {raw: false, header:1})
            aoa.forEach(function(r, i) {
                var cells = {}
                r.forEach(function(c, j) { cells[j] = ({ text: c }) })
                o.rows[i] = { cells: cells }
            })
            out.push(o)
        })
        return out
    }

    export default {
        name: 'Spreadsheet',
        props: [ 'item' ],
        data() {
            return {
                content: '',
                currentItem: '',
                lineNumbers: true,
            }
        },
        mounted() {

            this.currentItem = this.item
            let self = this
            api.downloadItem({
                path: this.item.path,
                responseType: 'arraybuffer'
            })
                .then((res) => {
                    let bytes = new Uint8Array(res)
                    let wb = XLSX.read(bytes, {type: 'array'})
                    self.ss = new Spreadsheet('#out-table', {mode: 'read', row: {len: 1500}, col: {len: 104}})
                        .loadData(stox(wb))
                })
                    .catch(error => this.handleError(error))
        },
        methods: {
        },
    }
</script>

<style scoped>
    @media (min-width: 1100px) {
        .modal-card {
            width: 100%;
            min-width: 640px;
        }
    }

    .preview {
        min-height: 450px;
    }
</style>
