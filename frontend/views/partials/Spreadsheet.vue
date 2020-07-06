<template>
  <div>
    <div class="modal-card">
      <header class="modal-card-head">
        <p class="modal-card-title">
          {{ currentItem.name }}
        </p>
      </header>
      <section class="modal-card-body preview">
        <div id="out-table" />
      </section>
      <footer class="modal-card-foot">
        <button v-if="can('write')" class="button" type="button" @click="saveFile()">
          {{ lang('Save') }}
        </button>
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

    function xtos(sdata) {
        var out = XLSX.utils.book_new()
        sdata.forEach(function(xws) {
            var aoa = [[]]
            var rowobj = xws.rows
            for(var ri = 0; ri < rowobj.len; ++ri) {
                var row = rowobj[ri]
                if(!row) continue
                aoa[ri] = []
                Object.keys(row.cells).forEach(function(k) {
                    var idx = +k
                    if(isNaN(idx)) return
                    aoa[ri][idx] = row.cells[k].text
                })
            }
            var ws = XLSX.utils.aoa_to_sheet(aoa)
            XLSX.utils.book_append_sheet(out, ws, xws.name)
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
                    self.ss = new Spreadsheet('#out-table')
                        .loadData(stox(wb))
                })
                    .catch(error => this.handleError(error))
        },
        methods: {
            saveFile() {
                var wb = xtos(this.ss.getData())
                var wopts = { bookType:'xlsx', bookSST:false, type:'array' }
                var wbout = XLSX.write(wb, wopts)

                api.saveContent({
                    name: this.item.name,
                    content: wbout,
                })
                    .then(() => {
                        this.$toast.open({
                            message: this.lang('Updated'),
                            type: 'is-success',
                        })
                        this.$parent.close()
                    })
                    .catch(error => this.handleError(error))
            }
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
